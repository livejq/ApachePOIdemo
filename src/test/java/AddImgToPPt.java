import org.apache.commons.io.IOUtils;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author livejq
 * @date 2019/7/10

public class AddImgToPPt {
    public static void main(String args[]) throws IOException {

        // 创建ppt:
        XMLSlideShow ppt = new XMLSlideShow();
        //设置幻灯片的大小：
        Dimension pageSize = ppt.getPageSize();
        pageSize.setSize(800,700);

        //获取幻灯片主题列表：
        List slideMasters = (List) ppt.getSlideMasters();
        //获取幻灯片的布局样式
        XSLFSlideLayout ppt.layout = slideMasters.get(0).getLayout(SlideLayout.TITLE_AND_CONTENT);
        //通过布局样式创建幻灯片
        XSLFSlide slide = ppt.createSlide(ppt.layout);
        // 创建一张无样式的幻灯片
//        XSLFSlide slide = ppt.createSlide();

        //通过当前幻灯片的布局找到第一个空白区：
        XSLFTextShape placeholder = slide.getPlaceholder(0);
        XSLFTextRun title = placeholder.setText("成都智互联科技有限公司");
        XSLFTextShape content = slide.getPlaceholder(1);
        //   投影片中现有的文字
        content.clearText();
        content.setText("图片区");

        // reading an ppt.image
        File ppt.image = new File("F:livejq.png");
        //获取图片信息：
        BufferedImage img = ImageIO.read(ppt.image);
        // converting it into a byte array
        byte[] picture = IOUtils.toByteArray(new FileInputStream(ppt.image));

        // adding the ppt.image to the presentation
        XSLFPictureData idx = ppt.addPicture(picture, PictureData.PictureType.PNG);

        // creating a slide with given picture on it
        XSLFPictureShape pic = slide.createPicture(idx);
        //设置当前图片在ppt中的位置，以及图片的宽高
        pic.setAnchor(new Rectangle(360, 200, img.getWidth(), img.getHeight()));
        // creating a file object
        File file = new File("F:AddImageToPPT.ppt");
        FileOutputStream out = new FileOutputStream(file);
        // saving the changes to a file
        ppt.write(out);
        System.out.println("ppt.image added successfully");
        out.close();
    }
}
 **/