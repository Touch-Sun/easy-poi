package com.touchsun.easypoi.ppt.business;

import org.apache.poi.sl.usermodel.PictureShape;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.io.IOException;

/**
 * 业务场景<br/>
 * 需求：替换PPT中的图片为新的图片<br/>
 * 限制：保证文字内容正确、保证排版格式正确、保证文件不出现打不开的情况<br/>
 * 设计：封装静态工具，保证接口简洁性<br/>
 * 兼容：pptx<br/>
 * 性能：优先考虑功能、其次考虑性能优化<br/>
 *
 * @author Lee
 */
public class ImageReplace {

    /**
     * 替换图片信息
     *
     * @param ppt PPT对象
     * @param pictureData 图片字节数组
     * @param index 第几个形状 索引从0开始
     */
    public static void replaceImage(XMLSlideShow ppt, byte[] pictureData, int index) throws IOException {
        int num = 0;
        for (XSLFSlide slide : ppt.getSlides()) {
            // 获取每一张幻灯片中的shape
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof PictureShape) {
                    if(num == index){
                        XSLFPictureShape pictureShape = (XSLFPictureShape) shape;
                        pictureShape.getPictureData().setData(pictureData);
                        break;
                    }
                    num++;
                }
            }
        }
    }

}
