package ppt.image;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class AddingImage {
    public static void main(String args[]) throws IOException {

        //creating a presentation
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(".\\temp\\addingImage.ppt"));

        //creating a slide in it
        XSLFSlide slide = ppt.createSlide();

        //reading an ppt.image
        File image = new File("F:\\livejq.png");

        //converting it into a byte array
        byte[] picture = IOUtils.toByteArray(new FileInputStream(image));

        //adding the ppt.image to the presentation
        XSLFPictureData idx = ppt.addPicture(picture, XSLFPictureData.PictureType.PNG);

        //creating a slide with given picture on it
        XSLFPictureShape pic = slide.createPicture(idx);

        //creating a file object
        File file = new File(".\\temp\\addingImage.ppt");
        FileOutputStream out = new FileOutputStream(file);

        //saving the changes to a file
        ppt.write(out);
        System.out.println("ppt.image added successfully");
        out.close();
    }
}
