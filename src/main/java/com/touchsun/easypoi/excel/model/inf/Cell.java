package com.touchsun.easypoi.excel.model.inf;


/**
 * Excel中的 单元格（Cell）接口
 *
 * @author WangXiang
 * @date 2022/7/14 15:57
 */
public interface Cell {

    /**
     * 设置此单元格的值
     * 实现需要判断具体类型进行处理
     *
     * @param value 写入单元格的具体数据
     */
    void setValue(Object value);

    /**
     * 获取此单元格的数据
     *
     * @return 单元格数据
     */
    Object getValue();

    /**
     * 获取当前单元格所在的行
     *
     * @return 此单元格的所在行
     */
    Row getRow();

    /**
     * 获取当前单元格所在的Sheet
     *
     * @return 此单元格的所在Sheet
     */
    Sheet getSheet();

    /**
     * 获取单元格所在列的索引
     *
     * @return Sheet中包含此单元格的列的索引（从 0 开始）
     */
    int getColumnIndex();

    /**
     * 获取单元格所在行的索引
     *
     * @return Sheet中包含此单元格的行的索引（从 0 开始）
     */
    int getRowIndex();

    /**
     * 置空此单元格数据（不清除格式）
     */
    void remove();

    /**
     * 完全清除此单元格（清除格式）
     */
    void clear();


}
