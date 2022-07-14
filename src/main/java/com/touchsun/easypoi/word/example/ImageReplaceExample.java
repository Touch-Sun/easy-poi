package com.touchsun.easypoi.word.example;


import com.touchsun.easypoi.word.business.ImageReplace;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Word图片替换案例
 *
 * @author Lee
 */
public class ImageReplaceExample {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        // 原Word文件地址
        String filePath = "D:\\backup\\2022\\0713\\POITest.docx";
        // 替换后新的Word文件地址
        String newFilePath = "D:\\backup\\2022\\0713\\new22POITest.docx";
        // 要替换的备用图片
        String imgPath = "D:\\backup\\2022\\0712\\bj.jpeg";
        // 1. 读取为XWPFDocument
        XWPFDocument docx = readDoc(filePath);
        // 2. 根据分析出来的XML中图片的ID替换为新的流
        ImageReplace.replacePicture("rId6", new FileInputStream(imgPath), docx);
        // 3. 保存文件
        OutputStream outputStream = new FileOutputStream(newFilePath);
        docx.write(outputStream);
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
