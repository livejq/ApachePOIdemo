package ppt.image;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;

/**
 * @author livejq
 * @date 2019/7/10
 **/
public class ReadingImage {
    public static void main(String args[]) throws IOException{

        //open an existing presentation
        File file = new File(".\\temp\\addingImage.ppt");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //reading all the pictures in the presentation
        for(XSLFPictureData data : ppt.getPictureData()){
            String fileName = data.getFileName();
            String extension = data.getType().extension;
            System.out.println("picture name: <" + fileName + ">");
            System.out.println("picture format: <" + extension + ">");
            System.out.println("=====================================");
            // 获取图片文件
            FileOutputStream fileOut = new FileOutputStream(new File(".\\temp\\" + data.getIndex() + extension));
            fileOut.write(data.getData());
        }
        System.out.println("报告总共" + ppt.getSlides().size() + "张幻灯片");

        //saving the changes to a file
        FileOutputStream out = new FileOutputStream(file);
        ppt.write(out);
        out.close();
    }
}
