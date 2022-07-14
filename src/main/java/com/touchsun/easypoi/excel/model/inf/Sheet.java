package com.touchsun.easypoi.excel.model.inf;


import java.util.List;

/**
 * Excel中的 工作表（Sheet）接口
 *
 * @author WangXiang
 * @date 2022/7/14 15:57
 */
public interface Sheet {

    /**
     * 获取当前 Sheet 所在的 Excel
     *
     * @return Excel对象
     */
    Excel getExcel();

    /**
     * 获取当前工作表的的索引
     *
     * @return 索引值（从 0 开始）
     */
    int getSheetIndex();

    /**
     * 获取当前工作表的名称
     *
     * @return 工作表名称
     */
    String getSheetName();

    /**
     * 根据行索引，获取指定行信息
     *
     * @param rowIndex 行索引（从 0 开始）
     * @return 行对象
     */
    Row getRow(int rowIndex);

    /**
     * 获取当前 Sheet 的全部行
     *
     * @return 行对象组成的 List
     */
    List<Row> getRowAll();

    /**
     * 在指定行索引上设置行对象
     *
     * @param rowIndex 行索引
     * @param row      行对象
     */
    void setRow(int rowIndex, Row row);

    /**
     * 删除指定行（保留格式）
     *
     * @param rowIndex 指定行索引
     */
    void removeRow(int rowIndex);

    /**
     * 清除指定行（不保留格式）
     *
     * @param rowIndex
     */
    void clearRow(int rowIndex);

    /**
     * 删除当前工作表的所有行（保留格式）
     */
    void remove();

    /**
     * 清除当前工作表的所有行（不保留格式）
     */
    void clear();

    /**
     * 灵活单元格
     * 将单元格信息强制写入指定行，列
     * @param row 行
     * @param col 列
     * @param cell 单元格信息
     */
    void setFlexibleCell(int row, int col, Cell cell);


}
