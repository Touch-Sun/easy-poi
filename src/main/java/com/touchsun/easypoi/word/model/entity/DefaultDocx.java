package com.touchsun.easypoi.word.model.entity;

import cn.hutool.core.util.StrUtil;
import com.touchsun.easypoi.word.business.ImageReplace;
import com.touchsun.easypoi.word.virtual.Docx;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import java.io.File;
import java.io.IOException;
import java.util.Map;

/**
 * 默认文档
 *
 * @author Lee
 */
public class DefaultDocx extends Docx {

    private XWPFDocument document;

    public DefaultDocx(String path) throws IllegalAccessException, InvalidFormatException, IOException {
        File file = new File(path);
        if (!file.exists()) {
            throw new IllegalAccessException(StrUtil.format("无法找到文件: {}", path));
        }
        this.document = new XWPFDocument(OPCPackage.open(file));
    }

    @Override
    public Map<String, XWPFPictureData> getAllPictureNodeMap() {
        return ImageReplace.getPicIdMap(this.document);
    }
}
