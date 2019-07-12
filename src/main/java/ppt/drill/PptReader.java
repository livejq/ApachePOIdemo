package ppt.drill;

import org.apache.poi.sl.usermodel.*;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * @author livejq
 * @date 2019/7/11
 **/
public class PptReader {

    private String fileName;

    public PptReader(){};

    public PptReader(String fileName) {
        this.fileName = fileName;
    }

    public boolean readPpt() throws IOException {

        if(fileName == null || fileName.equals("")) {
            return false;
        }
        // 读取ppt演示文档
        try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(fileName))) {
            // 获取ppt的一些属性（标题，创建者，最后修改时间等）
            System.out.println(ppt.getProperties().getCoreProperties().getTitle());
            System.out.println(ppt.getProperties().getCoreProperties().getCreator());
            System.out.println(ppt.getProperties().getCoreProperties().getModified());
            System.out.println(ppt.getProperties().getCoreProperties().getLastModifiedByUser());


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
            // 得先获取到该幻灯片的板式，了解了大致布局后再做解析（这里的板式同上）
            System.out.println("内容文本框中，文本的内容：<<" + slide.getPlaceholder(1).getText()+ ">>");
            System.out.println("xx, 文本其中的某个内容：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getRawText());
            System.out.println("xx, 文本的字体大小：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getFontSize());
            System.out.println("xx, 文本的字体样式：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getFontFamily());
            // 获取内容文本框第一段中第一个出现的超链接（段落中的所有超链接按序组成数组）
            System.out.println("xx, 文本的超链接地址：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getHyperlink().getAddress());
            // 判断文字格式
            System.out.println("xx, 文本的字体是否粗体：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).isBold());
            System.out.println("xx, 文本的字体是否有脚注：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).isSubscript());
            // .....是否存在下划线等
            Color textFillColor = slide.getPlaceholder(1).getFillColor();
            System.out.println("xx, 文本框填充颜色:(红：" + textFillColor.getRed() + "，绿：" + textFillColor.getGreen() + "，蓝:" + textFillColor.getBlue() + ")");

            System.out.println("xx, 段前间距：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getSpaceBefore());
            System.out.println("xx, 段后间距：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getSpaceAfter());
            System.out.println("xx, 行高：" + slide.getPlaceholder(1).getTextParagraphs().get(0).getLineSpacing());

            /** 图片对象 */
            System.out.println("=====================================");
            List<XSLFPictureData> pictureData = ppt.getPictureData();
            for(XSLFPictureData data : pictureData) {
                String fileName = data.getFileName();
                PictureData.PictureType pictureFormat = data.getType();
                Dimension imageDimensionInPixels = data.getImageDimensionInPixels();
                long imgSize = data.getData().length;
                System.out.println("图片名称: <" + fileName + ">");
                System.out.println("图片类型: <" + pictureFormat.contentType + ">");
                System.out.println("图片后缀: <" + pictureFormat.extension + ">");
                System.out.println("图片分辨率:<" + imageDimensionInPixels.getWidth() + " X " + imageDimensionInPixels.getHeight() + " px >");
                System.out.println("图片存储大小:" + imgSize/1024.00 + " KB");
                System.out.println("----");
                // 获取图片文件
                FileOutputStream fileOut = new FileOutputStream(new File(".\\temp\\" + data.getIndex() + pictureFormat.extension));
                fileOut.write(data.getData());
            }
            for(XSLFSlide xslfSlide : ppt.getSlides()) {
                List<XSLFShape> shapeList = xslfSlide.getShapes();
                for(XSLFShape shape : shapeList) {
                    if (shape instanceof XSLFPictureShape) {
                        XSLFPictureShape pictureShape = (XSLFPictureShape) shape;
                        System.out.println("超链接：" + pictureShape.getShapeId());
                    }
                }
            }


            /** 表格对象 */
            System.out.println("=====================================");


//            XSLFTable tbl = slide.createTable();
//            tbl.setAnchor(new Rectangle(50, 50, 450, 300));
//
//            int numColumns = 3;
//            int numRows = 5;
//            XSLFTableRow headerRow = tbl.addRow();
//            headerRow.setHeight(50);
//            // header
//            for (int i = 0; i < numColumns; i++) {
//                XSLFTableCell th = headerRow.addCell();
//                XSLFTextParagraph p = th.addNewTextParagraph();
//                p.setTextAlign(TextParagraph.TextAlign.CENTER);
//                XSLFTextRun r = p.addNewTextRun();
//                r.setText("Header " + (i + 1));
//                r.setBold(true);
//                r.setFontColor(Color.white);
//                th.setFillColor(new Color(79, 129, 189));
//                th.setBorderWidth(TableCell.BorderEdge.bottom, 2.0);
//                th.setBorderColor(TableCell.BorderEdge.bottom, Color.white);
//
//                tbl.setColumnWidth(i, 150);  // all columns are equally sized
//            }
//
//            // rows
//
//            for (int rownum = 0; rownum < numRows; rownum++) {
//                XSLFTableRow tr = tbl.addRow();
//                tr.setHeight(50);
//                // header
//                for (int i = 0; i < numColumns; i++) {
//                    XSLFTableCell cell = tr.addCell();
//                    XSLFTextParagraph p = cell.addNewTextParagraph();
//                    XSLFTextRun r = p.addNewTextRun();
//
//                    r.setText("Cell " + (i + 1));
//                    if (rownum % 2 == 0)
//                        cell.setFillColor(new Color(208, 216, 232));
//                    else
//                        cell.setFillColor(new Color(233, 247, 244));
//
//                }
//            }
//
//            try (FileOutputStream out = new FileOutputStream("temp/demo01.pptx")) {
//                ppt.write(out);
//            }
        }

        return true;
    }
//            slide.getPlaceholder(1).getTextParagraphs().get(0).getTextRuns().get(0).getFontColor();
//            System.out.println("xx, 文本的字体颜色：" + );

            /*// 创建一张无样式的幻灯片（首页）
            XSLFSlide slide = ppt.createSlide();
            // 背景
            slide.getBackground().setFillColor(new Color(55, 55, 122));
            // 标题
            XSLFTextBox title = slide.createTextBox();   //创建文本框
            title.setAnchor(new Rectangle2D.Double(400, 100, 250, 100));  //设置文本框的位置
            // 段落1
            XSLFTextParagraph titleFontP = title.addNewTextParagraph();    //创建一个段落
            XSLFTextRun titleTextRun = titleFontP.addNewTextRun();      //创建文本
            titleTextRun.setText("xxxx大学--发布公告");                  //设置文本类容
            titleTextRun.setFontSize(26.00);  //设置标题字号
            titleTextRun.setBold(true);    //设置成粗体
            System.out.println("段落内容：" + titleFontP.getText() + "，是否加粗：" + titleTextRun.isBold() + "，字体大小：" + titleTextRun.getFontSize());
            // 段落2
            XSLFTextParagraph titlePr = title.addNewTextParagraph();
            titlePr.setSpaceBefore(-20D);     // 设置与上一行的行距 :20D(正数代表正常行高的百分比)
            titlePr.setLeftMargin(35D);        // 设置段落开头的空格数
            titlePr.setBulletFont("宋体");
    //        titlePr.setBulletStyle("微软雅黑");
            titlePr.setBulletFontColor(new Color(255, 51, 0));
            titlePr.setLineSpacing(50D);
            System.out.println("字体：" + titlePr.getBulletFont()
                    + "，段落开头的空格数:" + titlePr.getLeftMargin()
                    + "，与上一行的行距：" + titlePr.getSpaceBefore()
                    + "，行高：" + titlePr.getLineSpacing());
            XSLFTextRun xslfTextRun = titlePr.addNewTextRun();
            xslfTextRun.setText("新生报到时间");
            xslfTextRun.setFontSize(26D);
            // 文本框1
            XSLFTextBox textBox = slide.createTextBox();
            textBox.setAnchor(new Rectangle2D.Double(30, 150, 300, 150));
            XSLFTextRun paragraph = textBox.addNewTextParagraph().addNewTextRun();
            paragraph.setText("xxx科技有限公司");
            paragraph.setBold(true);
            paragraph.setFontSize(30D);
            // 文本框2
            XSLFTextBox textCityBox = slide.createTextBox();
            textCityBox.setAnchor(new Rectangle2D.Double(440, 390, 250, 100));
            XSLFTextRun city = textCityBox.addNewTextParagraph().addNewTextRun();
            city.setText("广州");
            city.setFontSize(20D);
            // 文本框3
            XSLFTextBox textTimeBox = slide.createTextBox();
            textTimeBox.setAnchor(new Rectangle2D.Double(400, 420, 400, 100));
            XSLFTextRun time = textTimeBox.addNewTextParagraph().addNewTextRun();
            time.setText("2018年12月10日-2019年1月28日");
            time.setFontSize(20D);

            // 测试图片数据
            ArrayList<String> imgs = new ArrayList<String>();
            imgs.add("F:\\livejq.png");
            imgs.add("F:\\livejq.png");
            // 在2个幻灯片中分别插入2张图片
            int insertImg = 2;
            int slideSize = 2;
            if (imgs.size() >= insertImg) {
                for (int i = 0; i < slideSize; i++) {
                    // 创建一张幻灯片(最好读取一个现有的ppt文件)
                    XSLFSlide slidePicture = ppt.createSlide();
                    // 文本框1
                    XSLFTextBox projectNameBox = slidePicture.createTextBox();
                    projectNameBox.setAnchor(new Rectangle2D.Double(150, 100, 200, 200));
                    XSLFTextRun projectName = projectNameBox.addNewTextParagraph().addNewTextRun();
                    projectName.setText("xxx班级");
                    projectName.setBold(true);
                    projectName.setFontSize(20.00);
                    // 文本框2
                    XSLFTextBox projectInfoBox = slidePicture.createTextBox();
                    projectInfoBox.setAnchor(new Rectangle2D.Double(280, 100, 400, 200));
                    XSLFTextRun projectInfo = projectInfoBox.addNewTextParagraph().addNewTextRun();
                    projectInfo.setText("xx地址：" + "成都市锦江区水三接166号");
                    projectInfo.setFontSize(14.00);
                    XSLFTextRun projectType = projectInfoBox.addNewTextParagraph().addNewTextRun();
                    projectType.setText("社区属性：" + "商住楼");
                    projectType.setFontSize(14.00);
                    XSLFTextRun projectDdNum = projectInfoBox.addNewTextParagraph().addNewTextRun();
                    projectDdNum.setText("合同规定：" + "10");
                    projectDdNum.setFontSize(14.00);
                    XSLFTextRun projectPushNum = projectInfoBox.addNewTextParagraph().addNewTextRun();
                    projectPushNum.setText("实际发布：" + "8");
                    projectPushNum.setFontSize(14.00);
                    // 文本框3
                    XSLFTextBox pushPic = slidePicture.createTextBox();
                    pushPic.setAnchor(new Rectangle2D.Double(150, 210, 400, 100));
                    XSLFTextRun pushPicTxt = pushPic.addNewTextParagraph().addNewTextRun();
                    pushPicTxt.setText("发布实景图:");
                    pushPicTxt.setFontSize(14.00);

                    for (int k = 0; k < insertImg; k++){
                        byte[] picture2 = IOUtils.toByteArray(new FileInputStream(imgs.get(k)));
                        XSLFPictureData idx2 = ppt.addPicture(picture2, PictureData.PictureType.JPEG);
                        XSLFPictureShape pic2 = slidePicture.createPicture(idx2);
                        if(k == 0){
                            pic2.setAnchor(new Rectangle(150, 260, 200, 240));
                        }else if (k == 1){
                            pic2.setAnchor(new Rectangle(400, 260, 200, 240));
                        }
                    }
                }
            }
            System.out.println("ppt.image added successfully");*/



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
*/

}
