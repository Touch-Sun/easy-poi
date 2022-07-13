package com.touchsun.easypoi.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI去读取一个单元格,之后重写到文件中
 *
 * @author Lee
 */
public class WriteCellExample {
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
            cell.setCellValue("你好POI,我是被重写后的值");
            // 打开输出流，进行对Excel文件的重写
            try(OutputStream os = new FileOutputStream("readExample.xls")) {
                // workbook 还是原来的对象，此处直接复用进行数据的重写
                workbook.write(os);
            }
        } catch (IOException exception) {
            exception.printStackTrace();
        }
    }
}
