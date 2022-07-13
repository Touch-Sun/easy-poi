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
public class MergeCellExample {
    public static void main(String[] args) {
        // 实例化HSSFWorkBook
        Workbook workbook = new HSSFWorkbook();
        // 打开一个输出流，在一个特定的文件中
        try(OutputStream outputStream = new FileOutputStream("mergeCellExample.xls")) {
            // 声明一个Sheet
            Sheet sheet = workbook.createSheet("mergeCellSheet");
            // 创建单元格
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            // 设置单元格的值
            cell.setCellValue("一个会被合并的单元格");
            // 使用Sheet提供的合并API进行单元格的合并
            sheet.addMergedRegion(new CellRangeAddress(0, 4, 0, 4));
            // 写入Excel文件
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
