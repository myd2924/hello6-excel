package com.myd.hello6excel.dto;

import com.myd.hello6excel.util.ExcelField;
import lombok.Data;

import java.io.Serializable;
import java.math.BigDecimal;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/15 17:03
 * @Description:
 */
@Data
public class User implements Serializable{

    private static final long serialVersionUID = 1459442494741418918L;

    @ExcelField(headerName = "姓名",must = true)
    private String name;

    @ExcelField(headerName = "工作",must = true)
    private String job;

    @ExcelField(headerName = "薪水",must = true)
    private BigDecimal salery;

    @ExcelField(headerName = "年纪",must = true)
    private int age;

    @ExcelField(headerName = "地址")
    private String addres;
}
