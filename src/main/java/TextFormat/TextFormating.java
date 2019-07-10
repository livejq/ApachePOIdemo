package TextFormat;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class TextFormating {
    public static void main(String args[]) throws IOException{

        //creating an empty presentation
        XMLSlideShow ppt = new XMLSlideShow();

        //getting the slide master object
        XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);

        //select a layout from specified list
        XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

        //creating a slide with title and content layout
        XSLFSlide slide = ppt.createSlide(slideLayout);

        //selection of title place holder
        XSLFTextShape body = slide.getPlaceholder(1);

        //clear the existing text in the slide
        body.clearText();

        //adding new paragraph
        XSLFTextParagraph paragraph = body.addNewTextParagraph();

        //formatting line 1

        XSLFTextRun run1 = paragraph.addNewTextRun();
        run1.setText("This is a colored line");

        //setting color to the text
        run1.setFontColor(java.awt.Color.red);

        //setting font size to the text
        run1.setFontSize(24.00);

        //moving to the next line
        paragraph.addLineBreak();

        //formatting line 2

        XSLFTextRun run2 = paragraph.addNewTextRun();
        run2.setText("This is a bold line");
        run2.setFontColor(java.awt.Color.CYAN);

        //making the text bold
        run2.setBold(true);
        paragraph.addLineBreak();

        //formatting line 3

        XSLFTextRun run3 = paragraph.addNewTextRun();
        run3.setText(" This is a Strike line");
        run3.setFontSize(12.00);

        //making the text italic
        run3.setItalic(true);

        //strike through the text
        run3.setStrikethrough(true);
        paragraph.addLineBreak();

        //formatting line 4

        XSLFTextRun run4 = paragraph.addNewTextRun();
        run4.setText(" This an underlined line");
        run4.setUnderlined(true);

        //underlining the text
        paragraph.addLineBreak();

        //creating a file object
        File file=new File("TextFormat.ppt");

        FileOutputStream out = new FileOutputStream(file);

        //saving the changes to a file
        ppt.write(out);
        out.close();
    }
}
