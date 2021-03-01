package com.yuchumian.tools.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import lombok.experimental.UtilityClass;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.ArrayUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Objects;

/**
 * @author yuchumian 2021-03-01
 **/
@UtilityClass
public class ExcelUtils {
    private final static String XLS = "xls";
    private final static String XLSX = "xlsx";

    public <T> void writer(List<T> list, String path, Class<T> clazz) throws FileNotFoundException {
        writer(list, new File(path), clazz);
    }
    public <T> void writer(List<T> list, File file, Class<T> clazz) throws FileNotFoundException {
        writer(list,new FileOutputStream(file),clazz);
    }

    /**
     * 写excel文件
     * @param list 数据集合
     * @param out 传入文件流
     * @param clazz 数据类型
     * @param mergeColumnIndex 需要合并的列编号
     */
    public <T> void writer(List<T> list, OutputStream out, Class<T> clazz, int... mergeColumnIndex) {
        ExcelWriterBuilder writerBuilder = EasyExcel.write(out, clazz);
        if (ArrayUtils.isNotEmpty(mergeColumnIndex)) {
            writerBuilder.registerWriteHandler(new ExcelMergeStrategy(list.size(), mergeColumnIndex));
        }
        writerBuilder.sheet("sheet1").doWrite(list);
    }

    /**
     * {@link this#excelToHtml(HSSFWorkbook)}
     * @param excelPath excel 路径
     * @return html 文本
     * @throws IOException  HSSFWorkbook()、XSSFWorkbook()
     */
    public static String excelToHtml(String excelPath) throws
            IOException, TransformerException, ParserConfigurationException {
        if(StringUtils.isBlank(excelPath)){
            return "";
        }
        return excelToHtml(new File(excelPath));
    }

    /**
     *  {@link this#excelToHtml(HSSFWorkbook)}
     * @param file excel 文件
     * @return html 文本
     * @throws IOException  HSSFWorkbook()、XSSFWorkbook()
     */
    public static String excelToHtml(File file) throws
            IOException, TransformerException, ParserConfigurationException {
        if(Objects.isNull(file)){
            return "";
        }
        HSSFWorkbook excelBook = buildHssfWorkbook(file);
        if (excelBook == null) {
            return "";
        }
        return excelToHtml(excelBook);
    }


    /**
     * excel 转 html 文本
     * @param excelBook xls
     * @return html string
     * @throws ParserConfigurationException {@link DocumentBuilderFactory#newDocumentBuilder()}  DocumentBuilder
     * @throws TransformerException {@link Transformer#transform(Source, Result)}   Transformer
     */
    public static String excelToHtml(HSSFWorkbook excelBook) throws ParserConfigurationException, TransformerException {
        Document htmlDocument = buildDocument(excelBook);
        String content = "";
        try (ByteArrayOutputStream outStream = new ByteArrayOutputStream()){
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(outStream);
            Transformer serializer = buildTransformer();
            serializer.transform(domSource, streamResult);
            content = new String(outStream.toByteArray(), StandardCharsets.UTF_8);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return content;
    }

    private static HSSFWorkbook buildHssfWorkbook(File file) throws IOException {
        String fileType = getFileType(file);
        HSSFWorkbook excelBook = new HSSFWorkbook();
        if (!XLS.equals(fileType) && !XLSX.equals(fileType)) {
            return null;
        }
        if(XLS.equals(fileType)){
            excelBook = new HSSFWorkbook(new FileInputStream(file));
        }
        if(XLSX.equals(fileType)){
            xlsxToXls(file, excelBook);
        }
        return excelBook;
    }

    private static String getFileType(File file) {
        String fileName = file.getName();
        return FilenameUtils.getExtension(fileName).toLowerCase();
    }

    private static void xlsxToXls(File file, HSSFWorkbook excelBook) throws IOException {
        ExcelTransform xlsTransform = new ExcelTransform();
        XSSFWorkbook workbookOld = new XSSFWorkbook(new FileInputStream(file));
        xlsTransform.transformXlsx(workbookOld, excelBook);
    }

    private static Document buildDocument(HSSFWorkbook excelBook) throws ParserConfigurationException {
        ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        excelToHtmlConverter.setOutputColumnHeaders(false);
        excelToHtmlConverter.setOutputRowNumbers(false);
        excelToHtmlConverter.processWorkbook(excelBook);
        return excelToHtmlConverter.getDocument();
    }

    private static Transformer buildTransformer() throws TransformerConfigurationException {
        TransformerFactory tf = TransformerFactory.newInstance();
        tf.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        tf.setAttribute(XMLConstants.ACCESS_EXTERNAL_STYLESHEET, "");
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        return serializer;
    }

}
