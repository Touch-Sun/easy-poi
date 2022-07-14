package com.touchsun.easypoi.excel.model.inf;


import java.util.List;

/**
 * Excel中的 行（Row）接口
 *
 * @author WangXiang
 * @date 2022/7/14 15:56
 */
public interface Row {

    /**
     * 获取当前行所在的Sheet
     *
     * @return 所在Sheet对象
     */
    Sheet getSheet();

    /**
     * 设置此行中指定列的单元格
     *
     * @param col  列的索引
     * @param cell 单元格对象
     * @throws
     */
    void setCell(int col, Cell cell);

    /**
     * 将单元格追加到此行的末尾
     * 如果当前行不存在则会创建在索引为 0 的位置上
     *
     * @param cell 单元格对象
     * @throws
     */
    void appendCell(Cell cell);

    /**
     * 删除指定列的单元格（保留格式）
     *
     * @param col 指定列
     */
    void removeCell(int col);

    /**
     * 删除指定列的单元格（不保留格式）
     *
     * @param col 指定列
     */
    void clearCell(int col);

    /**
     * 删除当前列（保留格式）
     */
    void remove();

    /**
     * 删除当前列（不保留格式）
     */
    void clear();

    /**
     * 获取当前行的索引
     *
     * @return 当前行的索引（从 0 开始）
     */
    int getRowIndex();

    /**
     * 将一组数据，按顺序填充到当前行的单元格（从指定列开始）
     *
     * @param col  指定列
     * @param list list 字符串 数字 时间 类型的 List
     */
    void fillRow(int col, List<Object> list);

    /**
     * 将一组数据，按顺序填充到当前行的单元格（从第 0 列开始）
     *
     * @param list 字符串 数字 时间 类型的 List
     */
    void fillRow(List<Object> list);


}
