package com.myd.hello6excel.controller;

import com.myd.hello6excel.dto.User;
import com.myd.hello6excel.resolver.UserResolver;
import com.myd.hello6excel.resultcollection.UserResultCollector;
import com.myd.hello6excel.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.util.StreamUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.PostConstruct;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/12 17:24
 * @Description: 单纯的一个上传excel文件 模板类
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

    /**
     * postman 127.0.0.1:8086/upload/excelFile
     * @param file
     * @return
     */
    @RequestMapping("/excelFile")
    @ResponseBody
    public Map<String,Object> uploadExcel(@RequestParam(value = "excelFile")MultipartFile file){
        //加锁限制并发和请求
        //记录失败的信息 放入缓存 设置有效期 供用户下载
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

    /**
     * 下载本地文件  127.0.0.1:8086/upload/downloadTemplate
     * @param request
     * @param response
     */
    @GetMapping("/downloadTemplate")
    public void downloadTemplate(HttpServletRequest request, HttpServletResponse response) throws FileNotFoundException {
        //request
        System.out.println("资源请求完整路径:"+request.getRequestURL());
        System.out.println("资源请求部分路径:"+request.getRequestURI());
        System.out.println("返回请求行中的参数部分:"+request.getQueryString());
        System.out.println("发出请求客户机的IP:"+request.getRemoteAddr());
        System.out.println("发出请求客户机的端口:"+request.getRemotePort());
        System.out.println("发出请求客户机名称:"+request.getRemoteHost());
        System.out.println("返回web服务器的IP:"+request.getLocalAddr());
        System.out.println("返回web服务器主机名:"+request.getLocalName());
        System.out.println("返回客户机请求方式:"+request.getMethod());
        System.out.println("协议："+request.getScheme());

        //下载本地文件 文件的存放路径
        String fileName = "user-model.xls";
        InputStream inputStream = new FileInputStream("src\\main\\resources\\user-model.xls");
        response.setHeader("Content-Disposition","attachment; filename=" + fileName );
        try {
            StreamUtils.copy(inputStream,response.getOutputStream());
            response.getOutputStream().flush();
            response.getOutputStream().close();
        } catch (IOException e) {
            throw new RuntimeException("下载文件异常，请稍后重试", e);
        }


    }
}
