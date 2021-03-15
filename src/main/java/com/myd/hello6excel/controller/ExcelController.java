package com.myd.hello6excel.controller;

import com.myd.hello6excel.dto.User;
import com.myd.hello6excel.resolver.UserResolver;
import com.myd.hello6excel.resultcollection.UserResultCollector;
import com.myd.hello6excel.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.PostConstruct;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/12 17:24
 * @Description:
 */
@Controller
@RequestMapping("/upload")
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    private UserResolver<User> userUserResolver;

    @PostConstruct
    private void postConstruct(){
        userUserResolver = new UserResolver<>(User.class);
        userUserResolver.setSheetName("");
        userUserResolver.setMaxRowNum(200);
    }

    @RequestMapping("/excelFile")
    @ResponseBody
    public Map<String,Object> uploadExcel(@RequestParam(value = "excelFile")MultipartFile file){
        //加锁限制并发和请求
        Map result = new HashMap();
        UserResultCollector userResultCollector = new UserResultCollector();
        try {
            userUserResolver.parse(file.getInputStream(),userResultCollector);
        } catch (IOException e) {
            e.printStackTrace();
        }
        result.put("data",userResultCollector.getSuccessData());
        return result;
    }
}
