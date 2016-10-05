package com.github.izhangzhihao.OfficeProducer;

import org.junit.Test;

import java.util.HashMap;
import java.util.UUID;

import static com.github.izhangzhihao.OfficeProducer.DocxProducer.CreateEncryptDocxFromTemplate;

/**
 * DocxProducer测试类
 */
@SuppressWarnings("SpellCheckingInspection")
public class DocxProducerTest {
    @Test
    public void CreateEncryptDocxFromTemplateTest() throws Exception {
        String TemplatePath = "/Template/2.docx";
        HashMap<String, String> parameters = new HashMap<>();
        parameters.put("colour", "green");
        parameters.put("icecream", "chocolate");
        HashMap<String, String> imageParameters = new HashMap<>();
        String prefix = "D:/头像/";
        imageParameters.put("bookmark", prefix + "/33.png");
        CreateEncryptDocxFromTemplate(TemplatePath, parameters, imageParameters, "D:/Desktop/test.docx", UUID.randomUUID().toString());
    }
}
