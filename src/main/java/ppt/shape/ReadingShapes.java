package ppt.shape;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.xslf.usermodel.*;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class ReadingShapes {
    public static void main(String args[]) throws IOException {

        //creating a slideshow
        File file = new File("temp/shapes.ppt");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //get slides
        List<XSLFSlide> slide = ppt.getSlides();

        //getting the shapes in the presentation
        System.out.println("幻灯片总页数:" + slide.size() + ", Shapes in the presentation:");
        for (int i = 0; i < slide.size(); i++){

            List<XSLFShape> xslfShapeList = slide.get(i).getShapes();
            System.out.println("第 " + (i+1) +" 页");
            System.out.println("===================");
            for (int j = 0; j < xslfShapeList.size(); j++){

                XSLFShape sh = xslfShapeList.get(j);
//                System.out.println(sh.getShapeName());

                // shapes's anchor which defines the position of this shape in the slide
                if (sh instanceof PlaceableShape) {
                    java.awt.geom.Rectangle2D anchor = ((PlaceableShape)sh).getAnchor();
                    System.out.println("锚:" + anchor.getY());
                }
                if (sh instanceof XSLFTextBox) {
                    XSLFTextBox textBox = (XSLFTextBox) sh;
                    System.out.println("文本框:"+textBox.getText());
                }
                if (sh instanceof XSLFConnectorShape) {
                    XSLFConnectorShape line = (XSLFConnectorShape) sh;
                    System.out.println("行:" +line.getShapeName());
                    // work with Line
                } else if (sh instanceof XSLFTextShape) {
                    XSLFTextShape shape = (XSLFTextShape) sh;
                    System.out.println("文本:" + shape.getText());
                    // work with a shape that can hold text
                } else if (sh instanceof XSLFPictureShape) {
                    XSLFPictureShape shape = (XSLFPictureShape) sh;
                    System.out.println("图片:" + shape.getShapeName() + "，超链接：" + shape.getHyperlink());
                    // work with Picture
                }
                if(j + 1 != xslfShapeList.size())
                    continue;
                //name of the shape
                System.out.println(">>>总数：" + xslfShapeList.size());
            }
        }

        FileOutputStream out = new FileOutputStream(file);
        ppt.write(out);
        out.close();
    }
}
