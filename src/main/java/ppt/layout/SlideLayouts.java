package ppt.layout;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;

import java.io.IOException;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class SlideLayouts {
    public static void main(String[] args) throws IOException {

        //create an empty presentation
        XMLSlideShow ppt = new XMLSlideShow();
        System.out.println("Available slide layouts:");

        int num = 0;
        //getting the list of all slide masters
        for(XSLFSlideMaster master : ppt.getSlideMasters()){

            //getting the list of the layouts in each slide master
            for(XSLFSlideLayout layout : master.getSlideLayouts()){

                num++;
                //getting the list of available slides
                System.out.println(layout.getType());
            }
        }
        System.out.println("布局类型总数" + num);
    }
}
