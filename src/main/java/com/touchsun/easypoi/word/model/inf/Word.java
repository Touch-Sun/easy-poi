package com.touchsun.easypoi.word.model.inf;

import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import java.util.Map;

/**
 * 抽象Word对象
 *
 * @author Lee
 */
public interface Word {

    Map<String, XWPFPictureData> getAllPictureNodeMap();
}
