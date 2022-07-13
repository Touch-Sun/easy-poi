package com.touchsun.easypoi.excel.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI创建一个单元格
 *
 * @author Lee
 */
public class CellExample {
    public static void main(String[] args) {
        // 实例化HSSFWorkBook
        Workbook workbook = new HSSFWorkbook();
        // 打开一个输出流，在一个特定的文件中
        try(OutputStream outputStream = new FileOutputStream("cellExample.xls")) {
            // 声明一个Sheet
            Sheet sheet = workbook.createSheet("cellSheet");
            // 使用Sheet创建一个行 [第1索引位置行] [行的位置从0开始索引] <=> 第2行
            Row row = sheet.createRow(1);
            // 使用Row创建一个Cell [第3索引位置列] [列的位置从0开始索引] <=> 第4列
            Cell cell = row.createCell(3);
            // 给这个单元格赋值
            cell.setCellValue("Hi EasyPoi!");
            // 写入Excel
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
