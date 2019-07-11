package ppt.drill;


import org.junit.Test;

import java.io.IOException;

public class PptReaderTest {

    @Test
    public void readPpt() throws IOException {
        PptReader pptReader = new PptReader(".\\temp\\demo01.pptx");
        pptReader.readPpt();
    }
}
