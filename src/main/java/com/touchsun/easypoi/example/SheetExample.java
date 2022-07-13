package com.touchsun.easypoi.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Apache POI Sheet使用示例<br/>
 * 本示例描述了如何使用POI创建工作表
 *
 * @author Lee
 */
public class SheetExample {
    public static void main(String[] args) {
        // 实例化HSSFWorkBook
        Workbook workbook = new HSSFWorkbook();
        // 打开一个输出流，在一个特定的文件中
        try(OutputStream outputStream = new FileOutputStream("sheetExample.xls")) {
            // 声明一个Sheet1
            Sheet sheet1 = workbook.createSheet("sheet1");
            // 声明一个Sheet2
            Sheet sheet2 = workbook.createSheet("sheet2");
            // 写入Excel
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
