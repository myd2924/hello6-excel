package com.myd.hello6excel.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/12 17:41
 * @Description:
 */
public final class ExcelUtil {

    private static final String DEFAULT_DATE_FORMATE = "yyyy-MM-dd  HH:mm:ss";

    private ExcelUtil(){}

    /**
     * CellType 类型 值
         CELL_TYPE_NUMERIC 数值型 0
         CELL_TYPE_STRING 字符串型 1
         CELL_TYPE_FORMULA 公式型 2
         CELL_TYPE_BLANK 空值 3
         CELL_TYPE_BOOLEAN 布尔型 4
         CELL_TYPE_ERROR 错误 5
     * @param cell
     * @return
     */
    public static String getCellStringValue(Cell cell){
        if(null == cell){
            return "";
        }

        switch (cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)){
                    Date date = cell.getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat(DEFAULT_DATE_FORMATE);
                    return sdf.format(date);
                }
                return BigDecimal.valueOf(cell.getNumericCellValue()).setScale(4,BigDecimal.ROUND_DOWN).stripTrailingZeros().toPlainString();
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString().trim();
            case Cell.CELL_TYPE_FORMULA:
                FormulaEvaluator formulaEvaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                return getCellValue(formulaEvaluator.evaluate(cell));
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
                default:
                    return "";

        }
    }

    private static String getCellValue(CellValue cell) {
        if(null == cell){
            return "";
        }

        String cellValue = "";
        switch(cell.getCellType()){
            case Cell.CELL_TYPE_STRING:
                cellValue = cell.getStringValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                cellValue = String.valueOf(cell.getNumberValue());
                break;
            default :
                break;
        }
        return cellValue;
    }
}
