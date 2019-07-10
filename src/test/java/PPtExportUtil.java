import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;


/**
 * @author livejq
 * @date 2019/7/9

public class PPtExportUtil {
    //找到当前文件夹下面的所有图片文件
    private  List<File> ImgList = new ArrayList<File>();
    public List findAllImgFile(File file) throws IOException {
//        File file = new File("F:\\workroom\\img");
        File[] files = file.listFiles();
        for (File file1 : files) {
            if (file1.isDirectory()) {
                findAllImgFile(file1);
            } else if (ImageIO.read(file1) != null) {
                ImgList.add(file1);
            }
        }
        return ImgList;
    }

    public static XMLSlideShow exportPPt() throws IOException {
        // 创建ppt:
        XMLSlideShow ppt = new XMLSlideShow();
        //设置幻灯片的大小：
        Dimension pageSize = ppt.getPageSize();
        pageSize.setSize(975, 730);

        // 创建一张无样式的幻灯片（首页）
        XSLFSlide slide = ppt.createSlide();
        //标题
        XSLFTextBox title = slide.createTextBox();   //创建文本框
        title.setAnchor(new Rectangle2D.Double(400, 100, 250, 100));  //设置文本框的位置
        XSLFTextParagraph titleFontP = title.addNewTextParagraph();    //创建一个段落
        XSLFTextRun titleTextRun = titleFontP.addNewTextRun();      //创建文本
        titleTextRun.setText("成都肛肠医院--发布");                  //设置文本类容
        titleTextRun.setFontSize(26.00);  //设置标题字号
//        titleTextRun.setBold(true);    //设置成粗体
        XSLFTextParagraph titlePr = title.addNewTextParagraph();
        titlePr.setSpaceBefore(-20D);     // 设置与上一行的行距 :20D
        titlePr.setLeftMargin(35D);        //设置段落开头的空格数
        titlePr.setBulletFont("宋体");
        XSLFTextRun xslfTextRun = titlePr.addNewTextRun();
        xslfTextRun.setText("媒体监测报告");
        xslfTextRun.setFontSize(26.00);
        //公司
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(30, 150, 300, 150));
        XSLFTextRun paragraph = textBox.addNewTextParagraph().addNewTextRun();
        paragraph.setText("智互联科技有限公司");
        paragraph.setBold(true);
        paragraph.setFontSize(30.00);

//      城市
        XSLFTextBox textCityBox = slide.createTextBox();
        textCityBox.setAnchor(new Rectangle2D.Double(440, 390, 250, 100));
        XSLFTextRun city = textCityBox.addNewTextParagraph().addNewTextRun();
        city.setText("成都");
        city.setFontSize(20.00);
//     时间
        XSLFTextBox textTimeBox = slide.createTextBox();
        textTimeBox.setAnchor(new Rectangle2D.Double(400, 420, 400, 100));
        XSLFTextRun time = textTimeBox.addNewTextParagraph().addNewTextRun();
        time.setText("2018年12月10日-2019年1月28日");
        time.setFontSize(20.00);

//   插入图片到ppt中 、每页显示两张
        //测试图片数据
        ArrayList<String> imgs = new ArrayList<String>();
        imgs.add("F:\\livejq.png");
        imgs.add("F:\\livejq.png");
        imgs.add("F:\\livejq.png");
        imgs.add("F:\\livejq.png");
        //获取图片信息：
//      BufferedImage img = ImageIO.read(image);
        if (imgs.size() > 0) {
            for (int i = 0; i < imgs.size(); i++) {
                //创建一张幻灯片
                XSLFSlide slidePicture = ppt.createSlide();
                //项目名字
                XSLFTextBox projectNameBox = slidePicture.createTextBox();
                projectNameBox.setAnchor(new Rectangle2D.Double(150, 100, 200, 200));
                XSLFTextRun projectName = projectNameBox.addNewTextParagraph().addNewTextRun();
                projectName.setText("万科京城");
                projectName.setBold(true);
                projectName.setFontSize(20.00);
                //项目信息
                XSLFTextBox projectInfoBox = slidePicture.createTextBox();
                projectInfoBox.setAnchor(new Rectangle2D.Double(280, 100, 400, 200));
                XSLFTextRun projectInfo = projectInfoBox.addNewTextParagraph().addNewTextRun();
                projectInfo.setText("社区位置：" + "成都市锦江区水三接166号");
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
                //发布实景图
                XSLFTextBox pushPic = slidePicture.createTextBox();
                pushPic.setAnchor(new Rectangle2D.Double(150, 210, 400, 100));
                XSLFTextRun pushPicTxt = pushPic.addNewTextParagraph().addNewTextRun();
                pushPicTxt.setText("发布实景图:");
                pushPicTxt.setFontSize(14.00);

                //       插入图片 、每页显示两张图片:
                int h = 2;
                for (int k = 0;k<h;k++){
                    if(i<imgs.size()){
                        byte[] picture2 = IOUtils.toByteArray(new FileInputStream(imgs.get(i++)));
                        XSLFPictureData idx2 = ppt.addPicture(picture2, PictureData.PictureType.JPEG);
                        XSLFPictureShape pic2 = slidePicture.createPicture(idx2);
                        if(k==0){
                            pic2.setAnchor(new java.awt.Rectangle(150, 260, 200, 240));
                        }else if (k==1){
                            pic2.setAnchor(new java.awt.Rectangle(400, 260, 200, 240));
                        }
                    }
                }
                if(i>0){
                    i=i-1;
                }
            }
        }

        System.out.println("image added successfully");
        return ppt;
    }

    public static void ExportPPtModel() throws IOException {
        //读取模板ppt
        SlideShow ppt = new XMLSlideShow(new FileInputStream("F:a2.pptx"));
        //提取文本信息
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

    public static void main(String[] args) throws IOException {
        XMLSlideShow xmlSlideShow = PPtExportUtil.exportPPt();

        File ppt = new File("F:\\a.ppt");

        xmlSlideShow.write(new FileOutputStream(ppt));
    }
}
 **/