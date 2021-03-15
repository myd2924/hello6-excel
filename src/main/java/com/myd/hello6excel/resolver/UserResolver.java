package com.myd.hello6excel.resolver;

import com.myd.hello6excel.util.ExcelResolver;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/15 17:23
 * @Description:
 */
public class UserResolver<TExcelBean> extends ExcelResolver<TExcelBean> {

    public UserResolver(Class<TExcelBean> beanClass) {
        super(beanClass);
    }
}
