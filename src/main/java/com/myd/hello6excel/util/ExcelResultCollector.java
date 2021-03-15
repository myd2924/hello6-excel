package com.myd.hello6excel.util;

import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/12 18:00
 * @Description:
 */
public interface ExcelResultCollector<TExcelBean> {

    void onHeaderParse(List<ExcelResolver.ExcelHeader> headers);

    void onRowSuccess(Row row, List<String> rowValues,TExcelBean data);

    void onRowFailed(Row row,List<String> rowValues,Exception e);
}
