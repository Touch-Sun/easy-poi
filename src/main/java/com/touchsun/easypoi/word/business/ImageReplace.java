package com.touchsun.easypoi.word.business;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.InputStream;
import java.util.List;
import java.util.Objects;

/**
 * 业务场景<br/>
 * 需求：替换Word中的图片为新的图片<br/>
 * 限制：保证文字内容正确、保证排版格式正确、保证文件不出现打不开的情况<br/>
 * 设计：封装静态工具，保证接口简洁性<br/>
 * 兼容：优先doc,其次docx<br/>
 * 性能：优先考虑功能、其次考虑性能优化<br/>
 *
 * @author Lee
 */
public class ImageReplace extends Business {


    /**
     * 替换指定ID处的图片
     *
     * @param picId 图片在WordXML中的ID
     * @param pictureStream 新的图片输入流
     * @param docx 文档对象
     */
    public static void replacePicture(String picId, InputStream pictureStream, XWPFDocument docx)
            throws InvalidFormatException {
        List<XWPFParagraph> paragraphs = docx.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            if(CollectionUtils.isNotEmpty(runs)) {
                XWPFRun run = runs.get(0);
                Node node = run.getCTR().getDomNode();
                // drawing 一个绘画的图片
                Node drawingNode = getChildNode(node, "w:drawing");
                if (drawingNode == null) {
                    continue;
                }
                // 绘画图片的宽和高
                Node extentNode = getChildNode(drawingNode, "wp:extent");
                NamedNodeMap extentAttrs = Objects.requireNonNull(extentNode).getAttributes();
                // 绘画图片具体引用
                Node blipNode = getChildNode(drawingNode, "a:blip");
                NamedNodeMap blipAttrs = Objects.requireNonNull(blipNode).getAttributes();
                String rid = blipAttrs.getNamedItem("r:embed").getNodeValue();
                if(rid.equals(picId)) {
                    System.out.println("找到图片ID: " + picId);
                    // 获取图片信息
                    PackagePart part = docx.getPartById(rid);
                    // 删除此节点
                    node.removeChild(drawingNode);
                    // 此节点处添加图片
                    String newPicId = docx.addPictureData(pictureStream, Document.PICTURE_TYPE_PNG);
                    // 添加图片
                    addPictureToRun(run,newPicId,Document.PICTURE_TYPE_PNG,
                            Integer.parseInt(extentAttrs.getNamedItem("cx").getNodeValue()),
                            Integer.parseInt(extentAttrs.getNamedItem("cy").getNodeValue()));
                }
            }
        }
    }


    /**
     * 获取一个节点的子节点
     *
     * @param node 节点
     * @param nodeName 节点名称
     * @return 子节点
     */
    private static Node getChildNode(Node node, String nodeName) {
        if (!node.hasChildNodes()) {
            return null;
        }
        NodeList childNodes = node.getChildNodes();

        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            if (nodeName.equals(childNode.getNodeName())) {
                return childNode;
            }
            childNode = getChildNode(childNode, nodeName);
            if (childNode != null) {
                return childNode;
            }
        }
        return null;
    }

    /**
     * 添加图片到run
     *
     * @param run run对象
     * @param blipId blipId
     * @param id 节点ID
     * @param width 图片宽度
     * @param height 图片高度
     */
    private static void addPictureToRun(XWPFRun run,String blipId,int id,long width, long height){

        CTInline inline =run.getCTR().addNewDrawing().addNewInline();

        String picXml = "" +
                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "         <pic:nvPicPr>" +
                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
                "            <pic:cNvPicPr/>" +
                "         </pic:nvPicPr>" +
                "         <pic:blipFill>" +
                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "            <a:stretch>" +
                "               <a:fillRect/>" +
                "            </a:stretch>" +
                "         </pic:blipFill>" +
                "         <pic:spPr>" +
                "            <a:xfrm>" +
                "               <a:off x=\"0\" y=\"0\"/>" +
                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                "            </a:xfrm>" +
                "            <a:prstGeom prst=\"rect\">" +
                "               <a:avLst/>" +
                "            </a:prstGeom>" +
                "         </pic:spPr>" +
                "      </pic:pic>" +
                "   </a:graphicData>" +
                "</a:graphic>";

        XmlToken xmlToken = null;
        try {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch(XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");
    }
}
