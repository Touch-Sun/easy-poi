package com.touchsun.easypoi.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI去读取一个单元格
 *
 * @author Lee
 */
public class ReadCellExample {
    public static void main(String[] args) {
        // 打开一个Excel文件的输入流
        try(InputStream ins = new FileInputStream("readExample.xls")) {
            // 根据输入流去创建Workbook工厂
            Workbook workbook = new HSSFWorkbook(ins);
            // 获取第1个Sheet 对应的起始位置即为：0
            Sheet sheet = workbook.getSheetAt(0);
            // 获取到第一行 对应的其实行索引即为：0
            Row row = sheet.getRow(0);
            // 获取到第一列 对应的列索引即为：0
            Cell cell = row.getCell(0);
            // 输出单元格的值
            System.out.println("cell = " + cell);
        } catch (IOException exception) {
            exception.printStackTrace();
        }
    }
}
