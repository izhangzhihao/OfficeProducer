package com.github.izhangzhihao.OfficeProducer;


import lombok.Cleanup;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.ProtectDocument;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBException;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 创建、操作Docx的一系列方法
 */
@SuppressWarnings({"JavaDoc", "SpellCheckingInspection", "WeakerAccess", "unused"})
@Slf4j
public class DocxProducer {

    private static boolean DELETE_BOOKMARK = false;

    private static org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();

    /**
     * 创建Docx的主方法
     *
     * @param templatePath    模板docx路径
     * @param parameters      参数和值
     * @param imageParameters 书签和图片
     * @return
     */
    private static WordprocessingMLPackage CreateWordprocessingMLPackageFromTemplate(String templatePath,
                                                                                     HashMap<String, String> parameters,
                                                                                     HashMap<String, String> imageParameters)
            throws Exception {
        @Cleanup InputStream docxStream = DocxProducer.class.getResourceAsStream(templatePath);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(docxStream);
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        //第一步 替换字符参数
        replaceParameters(documentPart, parameters);

        //第二步 插入图片
        replaceBookMarkWithImage(wordMLPackage, documentPart, imageParameters);
        return wordMLPackage;
    }

    /**
     * 创建Docx并保存
     *
     * @param templatePath    模板docx路径
     * @param parameters      参数和值
     * @param imageParameters 书签和图片
     * @param savePath        保存docx的路径
     * @return
     */
    public static void CreateDocxFromTemplate(String templatePath,
                                              HashMap<String, String> parameters,
                                              HashMap<String, String> imageParameters,
                                              String savePath)
            throws Exception {
        WordprocessingMLPackage wordMLPackage = CreateWordprocessingMLPackageFromTemplate(templatePath, parameters, imageParameters);

        //保存
        saveDocx(wordMLPackage, savePath);
    }


    /**
     * 创建Docx并加密保存
     *
     * @param templatePath    模板docx路径
     * @param parameters      参数和值
     * @param imageParameters 书签和图片
     * @param savePath        保存docx的路径
     * @return
     */
    public static void CreateEncryptDocxFromTemplate(String templatePath,
                                                     HashMap<String, String> parameters,
                                                     HashMap<String, String> imageParameters,
                                                     String savePath,
                                                     String passWord)
            throws Exception {
        WordprocessingMLPackage wordMLPackage = CreateWordprocessingMLPackageFromTemplate(templatePath, parameters, imageParameters);

        //加密
        ProtectDocument protection = new ProtectDocument(wordMLPackage);
        protection.restrictEditing(STDocProtect.READ_ONLY, passWord);

        //保存
        saveDocx(wordMLPackage, savePath);
    }

    /**
     * 从Docx模板文件创建Docx然后转化为pdf
     *
     * @param templatePath    模板docx路径
     * @param parameters      参数和值
     * @param imageParameters 书签和图片
     * @param savePath        保存pdf的路径
     * @return
     */
    public static void CreatePDFFromDocxTemplate(String templatePath,
                                                 HashMap<String, String> parameters,
                                                 HashMap<String, String> imageParameters,
                                                 String savePath)
            throws Exception {
        WordprocessingMLPackage wordMLPackage = CreateWordprocessingMLPackageFromTemplate(templatePath, parameters, imageParameters);

        //转化成PDF
        convertDocxToPDF(wordMLPackage, savePath);

    }

    /**
     * 保存当前Docx文件
     */
    private static void saveDocx(WordprocessingMLPackage wordMLPackage,
                                 String savePath)
            throws FileNotFoundException, Docx4JException {
        Docx4J.save(wordMLPackage, new File(savePath), Docx4J.FLAG_SAVE_ZIP_FILE);
    }

    /**
     * 替换模板中的参数
     *
     * @param documentPart
     * @param parameters
     * @throws JAXBException
     * @throws Docx4JException
     */
    private static void replaceParameters(MainDocumentPart documentPart,
                                          HashMap<String, String> parameters)
            throws JAXBException, Docx4JException {
        documentPart.variableReplace(parameters);
    }

    /**
     * 替换书签为图片
     *
     * @param wordMLPackage
     * @param documentPart
     * @param imageParameters
     * @throws Exception
     */
    private static void replaceBookMarkWithImage(WordprocessingMLPackage wordMLPackage,
                                                 MainDocumentPart documentPart,
                                                 HashMap<String, String> imageParameters)
            throws Exception {
        Document wmlDoc = documentPart.getContents();
        Body body = wmlDoc.getBody();
        // 提取正文中所有段落
        List<Object> paragraphs = body.getContent();
        // 提取书签并创建书签的游标
        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
        new TraversalUtil(paragraphs, rt);

        // 遍历书签
        for (CTBookmark bm : rt.getStarts()) {
            String bookmarkName = bm.getName();
            String imagePath = imageParameters.get(bookmarkName);
            if (imagePath != null) {
                File imageFile = new File(imagePath);
                InputStream imageStream = new FileInputStream(imageFile);
                // 读入图片并转化为字节数组，因为docx4j只能字节数组的方式插入图片
                byte[] bytes = IOUtils.toByteArray(imageStream);
                // 创建一个行内图片
                BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
                // createImageInline函数的前四个参数我都没有找到具体啥意思
                // 最后一个是限制图片的宽度，缩放的依据
                Inline inline = imagePart.createImageInline(null, null, 0, 1, false, 800);
                // 获取该书签的父级段落
                P p = (P) (bm.getParent());
                ObjectFactory factory = new ObjectFactory();
                // R对象是匿名的复杂类型，然而我并不知道具体啥意思，估计这个要好好去看看ooxml才知道
                R run = factory.createR();
                // drawing理解为画布？
                Drawing drawing = factory.createDrawing();
                drawing.getAnchorOrInline().add(inline);
                run.getContent().add(drawing);
                p.getContent().add(run);
            }
        }
    }

    /**
     * docx文档转换为PDF
     *
     * @param wordMLPackage
     * @param pdfPath       PDF文档存储路径
     * @throws Exception
     */
    public static void convertDocxToPDF(WordprocessingMLPackage wordMLPackage,
                                        String pdfPath)
            throws Exception {
        //HashSet<String> features = new HashSet<>();
        //features.add(PP_PDF_APACHEFOP_DISABLE_PAGEBREAK_LIST_ITEM);
        //WordprocessingMLPackage process = Preprocess.process(wordMLPackage, features);

        FileOutputStream fileOutputStream = new FileOutputStream(pdfPath);
        Docx4J.toPDF(wordMLPackage, fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();

        /*FOSettings foSettings = Docx4J.createFOSettings();
        foSettings.setWmlPackage(wordMLPackage);
        Docx4J.toFO(foSettings, fileOutputStream, Docx4J.FLAG_EXPORT_PREFER_XSL);*/
    }


    /**
     * 将书签替换为文字
     *
     * @param documentPart
     * @param data
     * @throws Exception
     */
    private static void replaceBookmarkContents(MainDocumentPart documentPart, Map<DataFieldName, String> data) throws Exception {

        org.docx4j.wml.Document wmlDocumentEl = documentPart.getContents();
        Body body = wmlDocumentEl.getBody();
        List<Object> paragraphs = body.getContent();

        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
        new TraversalUtil(paragraphs, rt);

        for (CTBookmark bm : rt.getStarts()) {

            // do we have data for this one?
            if (bm.getName() == null) continue;
            String value = data.get(new DataFieldName(bm.getName()));
            if (value == null) continue;

            try {
                // Can't just remove the object from the parent,
                // since in the parent, it may be wrapped in a JAXBElement
                List<Object> theList = null;
                if (bm.getParent() instanceof P) {
                    System.out.println("OK!");
                    theList = ((ContentAccessor) (bm.getParent())).getContent();
                } else {
                    continue;
                }

                int rangeStart = -1;
                int rangeEnd = -1;
                int i = 0;
                for (Object ox : theList) {
                    Object listEntry = XmlUtils.unwrap(ox);
                    if (listEntry.equals(bm)) {
                        if (DELETE_BOOKMARK) {
                            rangeStart = i;
                        } else {
                            rangeStart = i + 1;
                        }
                    } else if (listEntry instanceof CTMarkupRange) {
                        if (((CTMarkupRange) listEntry).getId().equals(bm.getId())) {
                            if (DELETE_BOOKMARK) {
                                rangeEnd = i;
                            } else {
                                rangeEnd = i - 1;
                            }
                            break;
                        }
                    }
                    i++;
                }

                if (rangeStart > 0 && rangeEnd > rangeStart) {

                    // Delete the bookmark range
                    for (int j = rangeEnd; j >= rangeStart; j--) {
                        theList.remove(j);
                    }

                    // now add a run
                    org.docx4j.wml.R run = factory.createR();
                    org.docx4j.wml.Text t = factory.createText();
                    run.getContent().add(t);
                    t.setValue(value);

                    theList.add(rangeStart, run);
                }

            } catch (ClassCastException cce) {
                log.error(cce.getMessage(), cce);
            }
        }


    }

}
