package com.touchsun.easypoi.ppt.example;

import com.touchsun.easypoi.ppt.business.ImageReplace;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * ppt 图片替换案例
 *
 * @author Lee
 */
public class ImageReplaceExample {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        // 原Word文件地址
        String filePath = "D:\\backup\\2022\\0713\\pptTest.pptx";
        // 替换后新的Word文件地址
        String newFilePath = "D:\\backup\\2022\\0713\\newPPTTest.pptx";
        // 要替换的备用图片
        String imgPath = "D:\\backup\\2022\\0712\\bj.jpeg";
        // 读取PPT对象
        XMLSlideShow ppt = readPPT(filePath);
        // 替换某个索引处的图片
        ImageReplace.replaceImage(ppt, IOUtils.toByteArray(new FileInputStream(imgPath)), 0);
        // 保存文件
        OutputStream outputStream = new FileOutputStream(newFilePath);
        ppt.write(outputStream);
    }

    /**
     * 拿到文档结构
     *
     * @param path 原PPT地址
     * @return XWPFDocument对象
     * @throws IOException IO异常
     */
    public static XMLSlideShow readPPT(String path) throws IOException, InvalidFormatException {
        return new XMLSlideShow(OPCPackage.open(path));
    }
}
