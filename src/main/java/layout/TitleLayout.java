package layout;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
/**
 * @author livejq
 * @date 2019/7/10
 **/
public class TitleLayout {
    public static void main(String args[]) throws IOException{

        //creating presentation
        XMLSlideShow ppt = new XMLSlideShow();

        //getting the slide master object
        XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);

        //get the desired slide layout
        XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE);

        //creating a slide with title layout
        XSLFSlide slide1 = ppt.createSlide(titleLayout);

        //selecting the place holder in it
        XSLFTextShape title1 = slide1.getPlaceholder(0);

        //setting the title init
        title1.setText("Tutorials point");

        //create a file object
        File file = new File("Titlelayout.ppt");
        FileOutputStream out = new FileOutputStream(file);

        //save the changes in a PPt document
        ppt.write(out);
        System.out.println("slide cretated successfully");
        out.close();
    }
}
