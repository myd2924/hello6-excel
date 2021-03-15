package com.myd.hello6excel.resultcollection;

import com.myd.hello6excel.dto.User;
import com.myd.hello6excel.util.ExcelResolver;
import com.myd.hello6excel.util.ExcelResultCollector;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/15 17:32
 * @Description:
 */
public class UserResultCollector implements ExcelResultCollector<User>{

    private List<ExcelResolver.ExcelHeader> headers;

    @Getter
    private List<User> successData = new ArrayList<>();

    private List<User> failData = new ArrayList<>();

    @Override
    public void onHeaderParse(List<ExcelResolver.ExcelHeader> headers) {
        this.headers=headers;
    }

    @Override
    public void onRowSuccess(Row row, List<String> rowValues, User data) {
        this.successData.add(data);
    }

    @Override
    public void onRowFailed(Row row, List<String> rowValues, Exception e) {
    }
}
