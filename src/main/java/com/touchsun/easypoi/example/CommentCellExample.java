package com.touchsun.easypoi.example;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * Apache POI Cell使用示例<br/>
 * 本示例描述了如何使用POI去给一个单元格加上一个批注
 *
 * @author Lee
 */
public class CommentCellExample {
    public static void main(String[] args) {
        // 首先打开一个Excel的输出流
        try(OutputStream os = new FileOutputStream("commentExample.xls")) {
            // 构建一个Excel对象
            Workbook workbook = new HSSFWorkbook();
            // 我们创建一个Sheet
            Sheet sheet = workbook.createSheet("批注展示");
            // 创建一个HSSF族长，它代表Excel图形相关的所有的顶级容器
            HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
            // 创建一个承载批注的单元格
            Cell cell = sheet.createRow(0).createCell(0);
            // 给单元格设置一个值
            cell.setCellValue("一个批注测试案例");
            // 通过顶级容器族长去创建一个批注,并指定它的样式
            HSSFComment comment = patriarch.createComment(
                    new HSSFClientAnchor(0, 0, 0, 0, (short) 4, 2, (short) 6, 5)
            );
            // 给批注设置一个值,这里采用的是富文本对象来进行值的设置
            comment.setString(new HSSFRichTextString("一个对于单元格的批注!"));
            // 最后我们将单元格的批注应用在此处额单元格中
            cell.setCellComment(comment);
            // 将Excel内容写入到输出流到文件中
            workbook.write(os);
        } catch (IOException exception) {
            exception.printStackTrace();
        }
    }
}
