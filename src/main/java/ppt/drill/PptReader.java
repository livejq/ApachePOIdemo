package ppt.drill;

import org.apache.poi.sl.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTable;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableRow;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.List;

/**
 * @author livejq
 * @date 2019/7/11
 **/
public class PptReader {

    private String fileName;

    public PptReader() {
    }

    ;

    public PptReader(String fileName) {
        this.fileName = fileName;
    }

    public boolean readPpt() throws IOException {

        if (fileName == null || fileName.length() == 0) {
            return false;
        }
        // 读取ppt演示文档
        try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(fileName))) {
            // 获取ppt的一些属性（标题，创建者，最后修改时间等）
            /*System.out.println(ppt.getProperties().getCoreProperties().getTitle());
            System.out.println(ppt.getProperties().getCoreProperties().getCreator());
            System.out.println(ppt.getProperties().getCoreProperties().getModified());
            System.out.println(ppt.getProperties().getCoreProperties().getLastModifiedByUser());*/


            /** 幻灯片对象 */
            System.out.println("=====================================");
            // 获取幻灯片总数
            System.out.println("幻灯片总数：" + ppt.getSlides().size());
            // 获取第一张幻灯片
            XSLFSlide slide = ppt.getSlides().get(0);
            System.out.println("幻灯片的切换方式：");
            // 获取幻灯片尺寸
            Dimension pageSize = ppt.getPageSize();
            System.out.println("幻灯片的高度：" + pageSize.getHeight()
                    + ",幻灯片的宽度：" + pageSize.getWidth());
            // 获取第一张幻灯片的xxx（板式中未设置标题则为null）
            System.out.println("幻灯片的布局类型:" + slide.getSlideLayout().getType());
            System.out.println("幻灯片的标题:" + slide.getTitle());
            System.out.println("幻灯片编号:" + slide.getSlideNumber());
            System.out.println("幻灯片主题:" + slide.getTheme().getName());
            System.out.println("幻灯片背景颜色:(红:" + slide.getBackground().getFillColor().getRed() + "，绿:" + ppt.getSlides().get(0).getBackground().getFillColor().getGreen() + "，蓝:" + ppt.getSlides().get(0).getBackground().getFillColor().getBlue() + ")");
            System.out.println("背景透明度：" + slide.getBackground().getFillColor().getAlpha());

            /** 文本框对象 */
            System.out.println("=====================================");
/*
            DecimalFormat format = new DecimalFormat("##.#");
            // 得先获取到该幻灯片的板式，了解了大致布局后再做解析（这里的板式同上）
            System.out.println("内容文本框中，文本的内容：<<" + slide.getPlaceholder(1).getText() + ">>");
            System.out.println("xx, 某段内容：" + slide.getPlaceholder(1).getTextParagraphs().get(1).getText());
            System.out.println("xx, 段落总数" + slide.getPlaceholder(1).getTextBody().getParagraphs().size());
            System.out.println("xx, 某段内容中的某个内容：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(1).getRawText());
            // 获取内容文本框第一段中第一个出现的超链接（段落中的所有超链接按序组成数组）
            System.out.println("xx, 某段内容中的某个内容的超链接地址：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(1).getHyperlink().getAddress());
            System.out.println("xx, 文本的字体大小：" + format.format(slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getFontSize()));
            System.out.println("xx, 文本的字体样式：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getFontFamily());
            System.out.println("xx, 文本的字体颜色：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getFontColor());
            // 判断文字格式
            System.out.println("xx, 文本的字体是否粗体：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).isBold());
            System.out.println("xx, 文本的字体是否有脚注：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).isSubscript());
            // .....是否存在下划线等
            Color textFillColor = slide.getPlaceholder(1).getFillColor();
            System.out.println("xx, 文本框填充颜色:(红：" + textFillColor.getRed() + "，绿：" + textFillColor.getGreen() + "，蓝:" + textFillColor.getBlue() + ")");

            System.out.println("xx, 段前间距：" + slide.getPlaceholder(1).getTextParagraphs().get(1).getSpaceBefore());
            System.out.println("xx, 段后间距：" + slide.getPlaceholder(1).getTextParagraphs().get(1).getSpaceAfter());
            System.out.println("xx, 行距：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getLineSpacing());
            System.out.println("xx, 行高：" + slide.getPlaceholder(1).getTextHeight());*/

            System.out.println("xx, 缩进级别：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getIndentLevel());
            System.out.println("xx, 首行缩进：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getIndent());
            System.out.println("xx, 左外边距：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getLeftMargin());
            System.out.println("xx, 右外边距：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getRightMargin());
//            System.out.println("xx, 段前符号颜色：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getBulletFontColor());
            System.out.println("xx, 段前符号：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getBulletFont());
            System.out.println("xx, 段前符号大小：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getBulletFontSize());
            System.out.println("xx, 段落对齐方式：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getTextAlign());
            System.out.println("xx, 默认字体大小：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getDefaultFontSize());
            System.out.println("xx, 默认字体样式：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getDefaultFontFamily());
            System.out.println("xx, 默认Tab大小：" + slide.getPlaceholder(2).getTextParagraphs().get(0).getDefaultTabSize());

            /** 图片对象 */
            System.out.println("=====================================");
            /*List<XSLFPictureData> pictureData = ppt.getPictureData();
            for (XSLFPictureData data : pictureData) {
                String fileName = data.getFileName();
                PictureData.PictureType pictureFormat = data.getType();
                Dimension imageDimensionInPixels = data.getImageDimensionInPixels();
                long imgSize = data.getData().length;
                System.out.println("图片名称: <" + fileName + ">");
                System.out.println("图片类型: <" + pictureFormat.contentType + ">");
                System.out.println("图片后缀: <" + pictureFormat.extension + ">");
                System.out.println("图片分辨率:<" + imageDimensionInPixels.getWidth() + " X " + imageDimensionInPixels.getHeight() + " px >");
                System.out.println("图片存储大小:" + imgSize / 1024.00 + " KB");
                System.out.println("----");
                // 获取图片文件
                FileOutputStream fileOut = new FileOutputStream(new File(".\\temp\\" + data.getIndex() + pictureFormat.extension));
                fileOut.write(data.getData());
            }*/

            /** 表格对象 */
            System.out.println("=====================================");
            // 获取第二张幻灯片
            XSLFSlide slide2 = ppt.getSlides().get(1);
            List<XSLFShape> shapes = slide2.getShapes();
            for(XSLFShape part : shapes){
                if(part instanceof XSLFTable){
                    XSLFTable table = (XSLFTable) part;
                    CTTable ctt = table.getCTTable();
                    CTTableProperties tp = ctt.getTblPr();
                    // 表格行数
                    System.out.println(ctt.getTrList().size());
                    // 列数
                    System.out.println(ctt.getTrList().get(0).getTcList().size());
                    // cell属性数量
                    System.out.println(ctt.getTrList().get(0).getTcList().get(0).getTxBody().getPList().size());
                    // cell内容
                    System.out.println(ctt.getTrList().get(0).getTcList().get(0).getTxBody().getPList().get(0).getRList().get(0).getT());
                }
            }

            /** 表格对象 */
            System.out.println("=====================================");


        }

        return true;
    }
}

    /*public static void ExportPPtModel() throws IOException {
        // 读取模板ppt
        SlideShow ppt = new XMLSlideShow(new FileInputStream("F:a2.pptx"));
        // 提取文本信息
        List slides = (List) ppt.getSlides();
        //   SlideShow slideShow = copyPage(slides.get(1), ppt,2);
        for (XSLFSlide slide : slides) {
            List<XSLFShape> shapes = slide.getShapes();
            for(int i=0;i<shapes.size();i++){
                Rectangle2D anchor = shapes.get(i).getAnchor();
                if (shapes.get(i) instanceof XSLFTextBox) {
                    XSLFTextBox txShape = (XSLFTextBox) shapes.get(i);
                    if (txShape.getText().contains("{schemeName}")) {
                        // 替换文字内容.用TextRun获取替换的文本来设置样式
                        TextRun rt = txShape.setText(txShape.getText().replace("{schemeName}", "测试方案"));
                        rt.setFontColor(Color.BLACK);
                        rt.setFontSize(20.0);
                        rt.setBold(true);
                        rt.setFontFamily("微软雅黑");
                    }
                    else if (txShape.getText().contains("{time}")) {
                        TextRun textRun = txShape.setText(txShape.getText().replace("{time}", "2019-1-19"));
                        textRun.setFontColor(Color.BLACK);
                        textRun.setFontSize(20.0);
                        textRun.setFontFamily("微软雅黑");
                    }   else if (txShape.getText().contains("{projectAdd}")) {
                        TextRun textRun = txShape.setText(txShape.getText().replace("{projectAdd}", "成都市经江区"));
                        textRun.setFontColor(Color.BLACK);
                        textRun.setFontSize(16.0);
                        textRun.setFontFamily("微软雅黑");
                    } else if (txShape.getText().contains("{rzl}")) {
                        TextRun textRun = txShape.setText(txShape.getText().replace("{rzl}", "90%"));
                        textRun.setFontColor(Color.BLACK);
                        textRun.setFontSize(16.0);
                        textRun.setFontFamily("微软雅黑");
                    }
                    else if (txShape.getText().contains("{cg}")) {
                        TextRun textRun = txShape.setText(txShape.getText().replace("{cg}", "30"));
                        textRun.setFontColor(Color.BLACK);
                        textRun.setFontSize(16.0);
                        textRun.setFontFamily("微软雅黑");
                    }
                    else if (txShape.getText().contains("{mediaImg2}")) {
                        byte[] bytes = IOUtils.toByteArray(new FileInputStream(ResourceUtils.getFile("classpath:static/ceshi4.jpg")));
                        PictureData pictureData = ppt.addPicture(bytes, XSLFPictureData.PictureType.JPEG);
                        XSLFPictureShape picture = slide.createPicture(pictureData);
                        picture.setAnchor(anchor);
                    }
                    else if (txShape.getText().contains("{mediaImg1}")) {
                        byte[] bytes = IOUtils.toByteArray(new FileInputStream(ResourceUtils.getFile("classpath:static/ceshi4.jpg")));
                        PictureData pictureData = ppt.addPicture(bytes, XSLFPictureData.PictureType.JPEG);
                        XSLFPictureShape picture = slide.createPicture(pictureData);
                        picture.setAnchor(anchor);
                    }
                    else if(txShape.getText().contains("{projectImg}")){
                        byte[] bytes = IOUtils.toByteArray(new FileInputStream(ResourceUtils.getFile("classpath:static/ceshi5.jpg")));
                        PictureData pictureData = ppt.addPicture(bytes, XSLFPictureData.PictureType.JPEG);
                        XSLFPictureShape picture = slide.createPicture(pictureData);
                        picture.setAnchor(anchor);
                    }
                }
            }
        }
        OutputStream outputStreams = new FileOutputStream("F:\\test2.pptx");
        ppt.write(outputStreams);
    }


}
*/