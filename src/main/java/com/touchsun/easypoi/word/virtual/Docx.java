package com.touchsun.easypoi.word.virtual;

import com.touchsun.easypoi.word.model.inf.Word;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import java.util.Map;

/**
 * Word 2007 及以上
 *
 * @author Lee
 */
public abstract class Docx implements Word {

    @Override
    public Map<String, XWPFPictureData> getAllPictureNodeMap() {
        return null;
    }
}
