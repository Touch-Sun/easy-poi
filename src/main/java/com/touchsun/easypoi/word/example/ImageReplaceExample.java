package com.touchsun.easypoi.word.example;


import com.touchsun.easypoi.word.business.ImageReplace;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;

/**
 * Word图片替换案例
 *
 * @author Lee
 */
public class ImageReplaceExample {

    /** 原Word文件地址 */
    private static String filePath = "D:\\backup\\2022\\0713\\POITest.docx";
    /** 替换后新的Word文件地址 */
    private static String newFilePath = "D:\\backup\\2022\\0713\\new22POITest.docx";
    /** 要替换的备用图片 */
    private static String imgPath = "D:\\backup\\2022\\0712\\bj.jpeg";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        testReplaceImage();
        testGetAllIdToPicMap();
    }

    /**
     * 测试替换指定标识位置的图片并保存
     *
     * @throws IOException IOException
     * @throws InvalidFormatException InvalidFormatException
     */
    public static void testReplaceImage() throws IOException, InvalidFormatException {
        XWPFDocument docx = readDoc(filePath);
        ImageReplace.replacePicture("rId6", new FileInputStream(imgPath), docx);
        OutputStream outputStream = new FileOutputStream(newFilePath);
        docx.write(outputStream);
    }

    /**
     * 测试文档中图片标识与图片对象Map
     */
    public static void testGetAllIdToPicMap() throws IOException, InvalidFormatException {
        XWPFDocument docx = readDoc(filePath);
        Map<String, XWPFPictureData> allPicIds = ImageReplace.getPicIdMap(docx);
        System.out.println(allPicIds);
    }

    /**
     * 拿到文档结构
     *
     * @param path 原Word地址
     * @return XWPFDocument对象
     * @throws IOException IO异常
     */
    public static XWPFDocument readDoc(String path) throws IOException, InvalidFormatException {
        return new XWPFDocument(OPCPackage.open(path));
    }
}
