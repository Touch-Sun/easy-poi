package com.touchsun.easypoi.excel.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Apache POI WorkBook使用示例<br/>
 * 本示例描述了如何使用WorkBook进行工作
 *
 * @author Lee
 */
public class WorkBookExample {
    public static void main(String[] args) {
        // 实例化HSSFWorkBook
        Workbook workbook = new HSSFWorkbook();
        // 打开一个输出流，在一个特定的文件中
        try(OutputStream outputStream = new FileOutputStream("workbookExample.xls")) {
            // 写入Excel
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
