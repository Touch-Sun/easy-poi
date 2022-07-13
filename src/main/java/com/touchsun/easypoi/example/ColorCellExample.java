package com.touchsun.easypoi.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI创建一个含有边框的单元格
 *
 * @author Lee
 */
public class ColorCellExample {
    public static void main(String[] args) {
        // 实例化HSSFWorkBook
        Workbook workbook = new HSSFWorkbook();
        // 打开一个输出流，在一个特定的文件中
        try(OutputStream outputStream = new FileOutputStream("colorCellExample.xls")) {
            // 声明一个Sheet
            Sheet sheet = workbook.createSheet("CellStyleSheet");
            // 设置单元格列的默认宽度
            sheet.setDefaultColumnWidth(20);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            // 使用CellStyle对单元格的格式进修饰
            // CellStyle 由Workbook创建
            CellStyle cellStyle = workbook.createCellStyle();
            // 开始设置颜色
            cellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 设置单元格的值
            cell.setCellValue("一个带有颜色的单元格");
            // 设置单元格的样式
            cell.setCellStyle(cellStyle);
            // 写入Excel文件
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
