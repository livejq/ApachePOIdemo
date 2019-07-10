package manager;

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
public class SortSlide {
    public static void main(String args[]) throws IOException{

        //opening an existing presentation
        File file = new File("example1.ppt");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //get the slides
        List<XSLFSlide> slides = ppt.getSlides();

        //selecting the fourth slide
        XSLFSlide selectesdslide = slides.get(2);

        //bringing it to the top
        ppt.setSlideOrder(selectesdslide, 0);

        //creating an file object
        FileOutputStream out = new FileOutputStream(file);

        //saving the changes to a file
        ppt.write(out);
        out.close();
    }
}
