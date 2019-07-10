package Shape;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class ReadingShapes {
    public static void main(String args[]) throws IOException {

        //creating a slideshow
        File file = new File("shapes.ppt");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //get slides
        List<XSLFSlide> slide = ppt.getSlides();

        //getting the shapes in the presentation
        System.out.println("幻灯片总页数:" + slide.size() + ", Shapes in the presentation:");
        for (int i = 0; i < slide.size(); i++){

            List<XSLFShape> sh = slide.get(i).getShapes();
            System.out.println("第 " + (i+1) +" 页");
            System.out.println("===================");
            for (int j = 0; j < sh.size(); j++){

                System.out.println(sh.get(j).getShapeName());
                if(j+1 != sh.size())
                    continue;
                //name of the shape
                System.out.println(">>>总数：" + sh.size());
            }
        }

        FileOutputStream out = new FileOutputStream(file);
        ppt.write(out);
        out.close();
    }
}
