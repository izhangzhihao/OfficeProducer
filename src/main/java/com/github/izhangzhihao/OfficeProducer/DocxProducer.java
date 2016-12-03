package com.github.izhangzhihao.OfficeProducer;


import lombok.Cleanup;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.ProtectDocument;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.servlet.http.HttpServletResponse;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.net.URLEncoder;
import java.util.*;

import static com.github.izhangzhihao.OfficeProducer.FileUtils.copy;
import static com.github.izhangzhihao.OfficeProducer.FileUtils.inputStreamToFile;
import static com.github.izhangzhihao.OfficeProducer.ListUtils.isNullOrEmpty;

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
     * @param templatePath        模板docx路径
     * @param parameters          参数和值
     * @param paragraphParameters 段落参数
     * @param imageParameters     书签和图片
     * @return
     */
    private static WordprocessingMLPackage CreateWordprocessingMLPackageFromTemplate(String templatePath,
                                                                                     HashMap<String, String> parameters,
                                                                                     HashMap<String, String> paragraphParameters,
                                                                                     HashMap<String, String> imageParameters)
            throws Exception {
        @Cleanup InputStream docxStream = DocxProducer.class.getResourceAsStream(templatePath);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(docxStream);
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        //第一步 替换字符参数
        if (parameters != null) {
            replaceParameters(documentPart, parameters);
        }

        //第二步 替换段落
        if (paragraphParameters != null) {
            replaceParagraph(documentPart, paragraphParameters);
        }

        //第三步 插入图片
        if (imageParameters != null) {
            replaceBookMarkWithImage(wordMLPackage, documentPart, imageParameters);
        }
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
                                              HashMap<String, String> paragraphParameters,
                                              HashMap<String, String> imageParameters,
                                              String savePath)
            throws Exception {
        WordprocessingMLPackage wordMLPackage = CreateWordprocessingMLPackageFromTemplate(templatePath, parameters, paragraphParameters, imageParameters);

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
                                                     HashMap<String, String> paragraphParameters,
                                                     HashMap<String, String> imageParameters,
                                                     String savePath,
                                                     String passWord)
            throws Exception {
        WordprocessingMLPackage wordMLPackage = CreateWordprocessingMLPackageFromTemplate(templatePath, parameters, paragraphParameters, imageParameters);

        //加密
        ProtectDocument protection = new ProtectDocument(wordMLPackage);
        protection.restrictEditing(STDocProtect.READ_ONLY, passWord);

        //保存
        saveDocx(wordMLPackage, savePath);
    }

    /**
     * 创建Docx并加密，返回InputStream
     *
     * @param templatePath    模板docx路径
     * @param parameters      参数和值
     * @param imageParameters 书签和图片
     * @return
     */
    public static InputStream CreateEncryptDocxStreamFromTemplate(String templatePath,
                                                                  HashMap<String, String> parameters,
                                                                  HashMap<String, String> paragraphParameters,
                                                                  HashMap<String, String> imageParameters,
                                                                  String passWord)
            throws Exception {
        WordprocessingMLPackage wordMLPackage = CreateWordprocessingMLPackageFromTemplate(templatePath, parameters, paragraphParameters, imageParameters);

        //加密
        ProtectDocument protection = new ProtectDocument(wordMLPackage);
        protection.restrictEditing(STDocProtect.READ_ONLY, passWord);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        wordMLPackage.save(baos);

        ByteArrayDataSource bads =
                new ByteArrayDataSource(baos.toByteArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        return bads.getInputStream();

    }

    /**
     * 根据模板创建docx文档并放到response的outputstream中
     *
     * @param templatePath
     * @param parameter
     * @param fileName
     * @param response
     */
    public static void CreateEncryptDocxToResponseFromTemplate(String templatePath,
                                                               HashMap<String, String> parameter,
                                                               HashMap<String, String> paragraphParameters,
                                                               HashMap<String, String> imageParameters,
                                                               String fileName,
                                                               HttpServletResponse response) throws Exception {
        final InputStream inputStream = CreateEncryptDocxStreamFromTemplate(templatePath, parameter, paragraphParameters, imageParameters, UUID.randomUUID().toString());

        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.addHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "utf-8"));
        copy(inputStream, response.getOutputStream());
    }

    /**
     * 根据模板创建docx文档返回File
     *
     * @param templatePath
     * @param parameter
     * @param fileName
     */
    public static File CreateEncryptDocxFileFromTemplate(String templatePath,
                                                         HashMap<String, String> parameter,
                                                         HashMap<String, String> paragraphParameters,
                                                         HashMap<String, String> imageParameters,
                                                         String fileName) throws Exception {
        final InputStream inputStream = CreateEncryptDocxStreamFromTemplate(templatePath, parameter, paragraphParameters, imageParameters, UUID.randomUUID().toString());
        File file = new File(fileName);
        inputStreamToFile(inputStream, file);
        return file;
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
                                                 HashMap<String, String> paragraphParameters,
                                                 HashMap<String, String> imageParameters,
                                                 String savePath)
            throws Exception {
        WordprocessingMLPackage wordMLPackage = CreateWordprocessingMLPackageFromTemplate(templatePath, parameters, paragraphParameters, imageParameters);

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
     * 根据字符串参数替换段落
     *
     * @param documentPart
     * @param paragraphParameters
     */
    private static void replaceParagraph(MainDocumentPart documentPart, HashMap<String, String> paragraphParameters) throws JAXBException, Docx4JException {
        //List<Object> tables = getAllElementFromObject(documentPart, Tbl.class);
        /*for (Map.Entry<String, String> entries : paragraphParameters.entrySet()) {
            final Tbl table = getTemplateTable(tables, entries.getKey());
            final List<Object> allElementFromObject = getAllElementFromObject(table, P.class);
            final P p = (P) allElementFromObject.get(1);
            appendParaRContent(p, entries.getValue());
        }*/
        final List<Object> allElementFromObject = getAllElementFromObject(documentPart, P.class);
        //final P p = (P) allElementFromObject.get(22);

        for (Object paragraph : allElementFromObject) {
            final P para = (P) paragraph;
            if (!isNullOrEmpty(para.getContent())) {
                final List<Object> content = para.getContent();
                final String stringFromContent = getStringFromContent(content);
                final String s = paragraphParameters.get(stringFromContent);
                if (s != null) {
                    appendParaRContent(para, s);
                }
            }
        }
    }

    /**
     * 从Content中获得内容
     *
     * @param content
     * @return
     */
    public static String getStringFromContent(List<Object> content) {
        String temp = "";
        for (Object o : content) {
            if (o.getClass() == R.class) {
                final Object text = ((JAXBElement) ((R) o).getContent().get(0)).getValue();
                if (text.getClass() == Text.class) {
                    temp += ((Text) text).getValue();
                }
            }
        }
        return temp;
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
                                                 Map<String, String> imageParameters)
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
                // 新建一个Run
                R run = factory.createR();
                // drawing 画布
                Drawing drawing = factory.createDrawing();
                drawing.getAnchorOrInline()
                        .add(inline);
                run.getContent()
                        .add(drawing);
                p.getContent()
                        .add(run);
            }
        }
    }


    /**
     * 获取模板中的表格
     *
     * @param tables
     * @param templateKey
     * @return
     * @throws Docx4JException
     * @throws JAXBException
     */
    private static Tbl getTemplateTable(List<Object> tables, String templateKey) throws Docx4JException, JAXBException {
        for (Object tbl : tables) {
            List<?> textElements = getAllElementFromObject(tbl, Text.class);
            for (Object text : textElements) {
                Text textElement = (Text) text;
                if (textElement.getValue() != null && textElement.getValue().equals(templateKey))
                    return (Tbl) tbl;
            }
        }
        return null;
    }


    /**
     * @param content
     * @Description: 追加段落内容
     */
    public static void appendParaRContent(P p, String content) {
        List<?> texts = getAllElementFromObject(p, Text.class);
        if (texts.size() > 0) {
            Text textToReplace = (Text) texts.get(0);
            textToReplace.setValue("");
        }
        if (content != null) {
            R run = new R();
            p.getContent().add(run);
            String[] contentArr = content.split("\n");
            Text text = new Text();
            text.setSpace("preserve");
            text.setValue("    " + contentArr[0]);
            run.getContent().add(text);

            for (int i = 1, len = contentArr.length; i < len; i++) {
                Br br = new Br();
                run.getContent().add(br);// 换行
                text = new Text();
                text.setSpace("preserve");
                text.setValue("    " + contentArr[i]);
                run.getContent().add(text);
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
     * see <a href='http://blog.csdn.net/zhyh1986/article/details/8766628'></a>
     * 允许你针对一个特定的类来搜索指定元素以及它所有的孩子，例如，你可以用它获取文档中所有的表格、表格中所有的行以及其它类似的操作
     *
     * @param obj
     * @param toSearch
     * @return
     */
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }


    /**
     * 替换段落
     *
     * @param placeholder
     * @param textToAdd
     * @param template
     * @param addTo
     */
    private static void replaceParagraph(String placeholder, String textToAdd, WordprocessingMLPackage template, ContentAccessor addTo) {
        // 1. get the paragraph
        List<Object> paragraphs = getAllElementFromObject(template.getMainDocumentPart(), P.class);

        P toReplace = null;
        for (Object p : paragraphs) {
            List<Object> texts = getAllElementFromObject(p, Text.class);
            for (Object t : texts) {
                Text content = (Text) t;
                if (content.getValue().equals(placeholder)) {
                    toReplace = (P) p;
                    break;
                }
            }
        }

        // we now have the paragraph that contains our placeholder: toReplace
        // 2. split into seperate lines
        String as[] = StringUtils.splitPreserveAllTokens(textToAdd, '\n');

        for (String ptext : as) {
            // 3. copy the found paragraph to keep styling correct
            P copy = XmlUtils.deepCopy(toReplace);

            // replace the text elements from the copy
            List<?> texts = getAllElementFromObject(copy, Text.class);
            if (texts.size() > 0) {
                Text textToReplace = (Text) texts.get(0);
                textToReplace.setValue(ptext);
            }

            // add the paragraph to the document
            addTo.getContent().add(copy);
        }

        // 4. remove the original one
        ((ContentAccessor) toReplace.getParent()).getContent().remove(toReplace);

    }
}
