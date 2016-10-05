package com.github.izhangzhihao.OfficeProducer;

import org.junit.Test;

import java.util.HashMap;

import static com.github.izhangzhihao.OfficeProducer.DocProducer.CreateDocFromTemplate;

/**
 * DocProducer测试类
 */
public class DocProducerTest {
    @Test
    public void CreateDocFromTemplateTest() throws Exception {
        String templatePath = "/Template/2.doc";
        HashMap<String, String> parameters = new HashMap<>();
        parameters.put("colour", "green");
        parameters.put("icecream", "chocolate");
        //HashMap<String, String> imageParameters = new HashMap<>();
        //String prefix = "D:/头像/";
        //imageParameters.put("bookmark", prefix + "/33.png");
        CreateDocFromTemplate(templatePath, parameters, "D:/Desktop/test.doc");
    }
}
