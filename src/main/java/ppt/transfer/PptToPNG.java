package ppt.transfer;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class PptToPNG {
    public static void main(String args[]) throws IOException{

        //creating an empty presentation
        File file = new File("temp/addingImage.ppt");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //getting the dimensions and size of the slide
        Dimension pgSize = ppt.getPageSize();
        List<XSLFSlide> slide = ppt.getSlides();

        for (int i = 0; i < slide.size(); i++) {
            BufferedImage img = new BufferedImage(pgSize.width, pgSize.height, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = img.createGraphics();

            //clear the drawing area
            graphics.setPaint(Color.white);
            graphics.fill(new Rectangle2D.Float(0, 0, pgSize.width, pgSize.height));

            //render
            slide.get(i).draw(graphics);

            //creating an ppt.image file as output
            FileOutputStream out = new FileOutputStream("ppt_image"+ i +".png");
            javax.imageio.ImageIO.write(img, "png", out);
            ppt.write(out);

            System.out.println("Image successfully created");
            out.close();
        }
    }
}
