package com.myd.hello6excel.common;

import com.myd.hello6excel.handler.Handler;
import org.springframework.context.ApplicationContext;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;
import org.springframework.core.io.support.ResourcePatternResolver;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import javax.annotation.Resource;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author <a href="mailto:mayuanding@qianmi.com">OF3787-马元丁</a>
 * @version 0.1.0
 * @Date:2021/3/16 14:29
 * @Description:
 */
@Component
public class ProcessHandler {

    @Resource
    private ApplicationContext applicationContext;

    private final List<Handler> handlers = new ArrayList<>(8);

    private final Map<String,org.springframework.core.io.Resource> resourceMap = new HashMap<>(8);

    @PostConstruct
    public void registerEventHandler(){
        String[] beanNames = applicationContext.getBeanNamesForType(Handler.class);
        //注册处理器
        for(String bean : beanNames){
            handlers.add(applicationContext.getBean(bean,Handler.class));
        }

        //加载资源文件
        ResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
        try {
            org.springframework.core.io.Resource[] resources = resolver.getResources("classpath*:*.xls");
            for(org.springframework.core.io.Resource resource : resources){
                resourceMap.put(resource.getFilename(),resource);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
