package ppt.merge;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
/**
 * @author livejq
 * @date 2019/7/10
 **/
public class MergingMultiplePresentations {
    public static void main(String args[]) throws IOException{

        //creating empty presentation
        XMLSlideShow ppt = new XMLSlideShow();

        //taking the two presentations that are to be merged
        String file1 = "temp/presentation1.ppt";
        String file2 = "temp/presentation2.ppt";
        String[] inputs = {file1, file2};

        for(String arg : inputs){

            FileInputStream inputstream = new FileInputStream(arg);
            XMLSlideShow src = new XMLSlideShow(inputstream);

            for(XSLFSlide srcSlide : src.getSlides()){

                //merging the contents
                ppt.createSlide().importContent(srcSlide);
            }
        }

        String file3 = "temp/combinedPresentation.ppt";

        //creating the file object
        FileOutputStream out = new FileOutputStream(file3);

        // saving the changes to a file
        ppt.write(out);
        System.out.println("Merging done successfully");
        out.close();
    }
}
