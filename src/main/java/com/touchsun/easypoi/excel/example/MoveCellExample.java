package com.touchsun.easypoi.excel.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI去移动一行数据到另一行
 *
 * @author Lee
 */
public class MoveCellExample {
    public static void main(String[] args) {
        // 首选读取Excel拿到Sheet对象
        try(InputStream ins = new FileInputStream("moveExample.xls")) {
            Workbook workbook = new HSSFWorkbook(ins);
            // 拿到第一个Sheet
            Sheet sheet = workbook.getSheetAt(0);
            // 开始移动 几行开始 几行结束 移动几行
            sheet.shiftRows(4, 18 ,5);
            // 写入文件
            try(OutputStream os = new FileOutputStream("moveExample.xls")) {
                workbook.write(os);
            }
        } catch (IOException exception) {
            exception.printStackTrace();
        }
    }
}
