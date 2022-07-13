package com.touchsun.easypoi.excel.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI创建一个展示日期格式的单元格
 *
 * @author Lee
 */
public class DateFormatCellExample {
    public static void main(String[] args) {
        // 实例化HSSFWorkBook
        Workbook workbook = new HSSFWorkbook();
        // 需要使用创建助手来设置单元格的格式
        CreationHelper creationHelper = workbook.getCreationHelper();
        // 打开一个输出流，在一个特定的文件中
        try(OutputStream outputStream = new FileOutputStream("dateFormatCellExample.xls")) {
            // 声明一个Sheet
            Sheet sheet = workbook.createSheet("DateSheet");
            // 设置单元格列的默认宽度
            sheet.setDefaultColumnWidth(20);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            // 使用CellStyle对单元格的格式进修饰
            // CellStyle 由Workbook创建
            CellStyle cellStyle = workbook.createCellStyle();
            // 设置这个CellStyle的单元格格式
            cellStyle.setDataFormat(
                    creationHelper.createDataFormat().getFormat("yyyy-MM-dd")
            );
            // 设置单元格的值
            cell.setCellValue(new Date());
            // 设置单元格的样式
            cell.setCellStyle(cellStyle);
            // 写入Excel文件
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
