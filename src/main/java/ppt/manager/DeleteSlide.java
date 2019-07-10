package ppt.manager;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class DeleteSlide {
    public static void main(String args[]) throws IOException {

        //Opening an existing slide
        File file = new File(".\\temp\\delete.ppt");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //deleting a slide
        ppt.removeSlide(2);

        System.out.println("=============delete successfully");

        //creating a file object
        FileOutputStream out = new FileOutputStream(file);

        //Saving the changes to the presentation
        ppt.write(out);
        out.close();
    }
}
