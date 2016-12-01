package com.github.izhangzhihao.OfficeProducer;

import org.docx4j.model.fields.merge.DataFieldName;
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
        String templatePath = "/Template/10.docx";
        HashMap<String, String> parameters = new HashMap<>();
        parameters.put("colour", "绿色");
        parameters.put("icecream", "巧克力");
        HashMap<String, String> imageParameters = new HashMap<>();
        String prefix = "D:/头像/";
        imageParameters.put("bookmark", prefix + "/33.png");

        HashMap<DataFieldName, String> map = new HashMap<>();
        map.put(new DataFieldName("projectName"), "校级项目");

        HashMap<String, String> paragraphParameters = new HashMap<>();
        paragraphParameters.put("test", "第三方公司的发生地方\n第三方风格的鬼地方个地方\n规划法国恢复规划法规\nsdfsf电饭锅电饭锅电饭锅地方个dfs");


        CreateEncryptDocxFromTemplate(templatePath, parameters, paragraphParameters, imageParameters, "D:/Desktop/test.docx", UUID.randomUUID().toString());
    }
}
