package com.touchsun.easypoi.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI合并几个单元格
 *
 * @author Lee
 */
public class FontCellExample {
    public static void main(String[] args) {
        // 实例化HSSFWorkBook
        Workbook workbook = new HSSFWorkbook();
        // 打开一个输出流，在一个特定的文件中
        try(OutputStream outputStream = new FileOutputStream("fontCellExample.xls")) {
            // 声明一个Sheet
            Sheet sheet = workbook.createSheet("fontCellSheet");
            // 创建单元格
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            // 获取样式对象
            CellStyle cellStyle = workbook.createCellStyle();
            // 获取字体对象
            Font font = workbook.createFont();
            // 设置字体样式
            font.setFontHeightInPoints((short)11);
            font.setFontName("Courier New");
            font.setItalic(true);
            font.setStrikeout(true);
            // 样式应用这个字体
            cellStyle.setFont(font);
            // 设置单元格的值
            cell.setCellValue("展示一个单元格的字体");
            // 单元格应用这个样式
            cell.setCellStyle(cellStyle);
            // 写入Excel文件
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
