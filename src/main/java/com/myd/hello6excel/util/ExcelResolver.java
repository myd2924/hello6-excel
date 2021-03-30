package com.myd.hello6excel.util;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.core.annotation.AnnotationUtils;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.util.*;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/15 09:52
 * @Description: 轻量级的解析，大数据量的还是走任务中心
 */
public class ExcelResolver<TExcelBean> {

    /**
     * excel bean 定义
     */
    protected ExcelBeanDefine excelBeanDefine;

    /**
     * 业务处理表单 默认取第一个
     */
    @Setter
    protected String sheetName;

    /**
     * 最大行数，如果小于等于0则不校验
     */
    @Setter
    protected int maxRowNum = -1;

    public ExcelResolver(Class<TExcelBean> beanClass){
        this.excelBeanDefine = new ExcelBeanDefine(beanClass);
    }

    /**
     * 解析
     * @param fileInputStream
     * @param collector
     */
    public void parse(InputStream fileInputStream,ExcelResultCollector<TExcelBean> collector){
        try(InputStream inputStream = fileInputStream){
            Workbook workbook = WorkbookFactory.create(inputStream);
            doParse(workbook,collector);
        } catch (IOException e) {
            throw new RuntimeException("文件读取失败，检查重新上传");
        } catch (InvalidFormatException e) {
            throw new RuntimeException("文件格式错误，请检查重新上传");
        }
    }

    /**
     * 解析
     * @param workbook
     * @param collector
     */
    public void doParse(Workbook workbook, ExcelResultCollector<TExcelBean> collector) {
        Sheet sheet = getSheet(workbook);
        List<ExcelHeader> headerList = null;

        Iterator<Row> rowIterator = sheet.rowIterator();
        int headerRowNum = -1;

        //找到表头
        while(rowIterator.hasNext()){
            Row headRow = rowIterator.next();
            if(null != headRow){
                headerList = excelBeanDefine.parseHeaderRow(headRow);
                if(CollectionUtils.isNotEmpty(headerList)){
                    collector.onHeaderParse(headerList);
                    headerRowNum = headRow.getRowNum();
                    break;
                }
            }
        }

        checkRowNum(sheet,headerRowNum);

        //处理表格内容
        while(rowIterator.hasNext()){
            Row dataRow = rowIterator.next();
            if(dataRow != null){
                parseDataRow(dataRow,headerList,collector);
            }
        }
    }

    /**
     * 解析表数据行
     * @param dataRow
     * @param headerList
     * @param collector
     */
    private void parseDataRow(Row dataRow, List<ExcelHeader> headerList, ExcelResultCollector<TExcelBean> collector) {
        //采集行数据
        List<String> rowValues = collectRowValues(dataRow,headerList);
        if(isBlankRow(rowValues)){
            //跳过空行
            return ;
        }

        try {
            TExcelBean data = excelBeanDefine.getBeanClass().newInstance();
            setDataField(data,headerList,rowValues);
            collector.onRowSuccess(dataRow,rowValues,data);
        } catch (Exception e) {
            collector.onRowFailed(dataRow,rowValues,e);
        }
    }

    /**
     * 设置数据字段
     * @param data
     * @param headerList
     * @param rowValues
     */
    private void setDataField(TExcelBean data, List<ExcelHeader> headerList, List<String> rowValues) {
        for(int i=0;i<headerList.size();i++){
            ExcelHeader header = headerList.get(i);
            String value = rowValues.get(i);

            if(header.isMust() && StringUtils.isEmpty(value)){
                throw new RuntimeException(String.format("列【%s】不能为空",header.getHeaderName()));
            }

            try{
                Field field = header.getField();
                if(String.class.equals(field.getType())){
                    field.set(data,value);
                } else if(BigDecimal.class.equals(field.getType())){
                    field.set(data,StringUtils.isNotEmpty(value)? new BigDecimal(value).stripTrailingZeros() : null);
                }
            } catch (Exception e){
                throw new RuntimeException(String.format("列【%s】数据格式错误",header.getHeaderName()));
            }
        }
    }

    private boolean isBlankRow(List<String> rowValues) {
        return CollectionUtils.isEmpty(rowValues) || rowValues.stream().allMatch(va -> StringUtils.isEmpty(va));
    }

    /**
     * 采集数据行
     * @param dataRow
     * @param headerList
     * @return
     */
    private List<String> collectRowValues(Row dataRow, List<ExcelHeader> headerList) {
        List<String> rowValue = new ArrayList<>();
        for(ExcelResolver.ExcelHeader header : headerList){
            Cell cell = dataRow.getCell(header.getColnumIndex());
            rowValue.add(ExcelUtil.getCellStringValue(cell));
        }
        return rowValue;
    }

    /**
     * 校验表格行
     * @param sheet
     * @param headerRowNum
     */
    private void checkRowNum(Sheet sheet, int headerRowNum) {
        if(headerRowNum < 0){
            throw new RuntimeException("excel表格没有表头");
        }
        if(sheet.getPhysicalNumberOfRows()-1-headerRowNum <= 0){
            throw new RuntimeException("excel表格没有内容");
        }
        if(maxRowNum > 0){
            if(sheet.getPhysicalNumberOfRows() -1 - headerRowNum > maxRowNum){
                throw new RuntimeException(String.format("excel表格行数不要超过%s条",maxRowNum));
            }
        }
    }

    /**
     * 获取表格
     * @param workbook
     * @return
     */
    private Sheet getSheet(Workbook workbook) {
        Sheet sheet = StringUtils.isNotEmpty(sheetName) ? workbook.getSheet(sheetName) : workbook.getSheetAt(0);
        if(null == sheet){
            throw new RuntimeException("没有需要处理的excel表格");
        }

        if(sheet.getPhysicalNumberOfRows() <= 0){
            throw new RuntimeException("excel表格没有内容");
        }
        return sheet;
    }

    @Data
    @Accessors(chain = true)
    public static class ExcelHeader{

        private String headerName;

        private boolean must;

        private int colnumIndex;

        private Field field;
    }

    /**
     * excel bean 定义
     */
    protected class ExcelBeanDefine{

        /**
         * excel bean类
         */
        @Getter
        private Class<TExcelBean> beanClass;

        /**
         * excel列名 -> excel bean字段
         */
        private Map<String,Field> headerNameFieldMap = new HashMap<>();

        /**
         * 必填的列名
         */
        private Set<String> mustHeaderNameSet = new HashSet<>();

        public ExcelBeanDefine(Class<TExcelBean> beanClass){
            this.beanClass = beanClass;
            Field[] fields = beanClass.getDeclaredFields();
            for(Field field : fields){
                if(Modifier.isStatic(field.getModifiers())){
                    continue;
                }
                ExcelField annotation = AnnotationUtils.findAnnotation(field, ExcelField.class);
                if(null != annotation && StringUtils.isNotEmpty(annotation.headerName())){
                    field.setAccessible(true);
                    headerNameFieldMap.put(annotation.headerName(),field);
                    if(annotation.must()){
                        mustHeaderNameSet.add(annotation.headerName());
                    }
                }
            }
        }

        /**
         * 解析表头
         */
        public List<ExcelHeader> parseHeaderRow(Row headerRow){
            List<ExcelHeader> headerList = new ArrayList<>();
            Set<String> existHeaderNameSet = new HashSet<>();
            //遍历表头
            Iterator<Cell> cellIterator = headerRow.cellIterator();
            while(cellIterator.hasNext()){
                Cell headerCell = cellIterator.next();
                String headerName = trimHeaderName(ExcelUtil.getCellStringValue(headerCell));
                Field field = headerNameFieldMap.get(headerName);
                if(null != field){
                    ExcelHeader excelHeader = new ExcelHeader()
                            .setHeaderName(headerName)
                            .setMust(mustHeaderNameSet.contains(headerName))
                            .setColnumIndex(headerCell.getColumnIndex())
                            .setField(field);
                    headerList.add(excelHeader);
                    existHeaderNameSet.add(headerName);
                }
            }

            //检查是否有遗漏列
            if(CollectionUtils.isNotEmpty(headerList)){
                mustHeaderNameSet.forEach(headerName->{
                    if(!existHeaderNameSet.contains(headerName)){
                        throw  new RuntimeException(String.format("excel缺少列【%s】",headerName));
                    }
                });
            }
            return headerList;
        }

        private String trimHeaderName(String cellStringValue) {
            if(cellStringValue.startsWith("*")){
                cellStringValue = cellStringValue.substring(1,cellStringValue.length());
            }

            if(cellStringValue.endsWith("*")){
                cellStringValue = cellStringValue.substring(0,cellStringValue.length()-1);
            }

            return cellStringValue;
        }


    }

}
