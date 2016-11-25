package com.github.izhangzhihao.OfficeProducer;

import lombok.Cleanup;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * 创建、操作Doc的一系列方法
 */
@SuppressWarnings({"JavaDoc", "WeakerAccess"})
public class DocProducer {

    /**
     * 创建Doc并保存
     *
     * @param templatePath 模板doc路径
     * @param parameters   参数和值
     *                     //* @param imageParameters 书签和图片
     * @param savePath     保存doc的路径
     * @return
     */
    public static void CreateDocFromTemplate(String templatePath,
                                             HashMap<String, String> parameters,
                                             //HashMap<String, String> imageParameters,
                                             String savePath)
            throws Exception {
        @Cleanup InputStream is = DocProducer.class.getResourceAsStream(templatePath);
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();

        //把range范围内的${}替换
        for (Map.Entry<String, String> next : parameters.entrySet()) {
            range.replaceText("${" + next.getKey() + "}",
                    next.getValue()
            );
        }

        @Cleanup OutputStream os = new FileOutputStream(savePath);
        //把doc输出到输出流中
        doc.write(os);
    }
}
