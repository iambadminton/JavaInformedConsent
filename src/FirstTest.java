import word.api.interfaces.IDocument;
import word.w2004.Document2004;
import word.w2004.elements.BreakLine;
import word.w2004.elements.Heading1;
import word.w2004.elements.Paragraph;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Created by a.shipulin on 29.06.18.
 */
public class FirstTest {
    public static void main(String[] args) {
        IDocument myDoc = new Document2004();
        myDoc.addEle(Heading1.with("Heading01").create());
        myDoc.addEle(BreakLine.times(2).create()); // two break lines
        myDoc.addEle(Paragraph.with("This document is an example of paragraph").create());

        String myWord = myDoc.getContent();


        try {
            File file = new File("c:\\29062018\\FirstJavaWordDoc.doc");
            String xmlTemplate = "";

            OutputStream outputStream = new FileOutputStream(file);
            //PrintWriter writer = new PrintWriter(new OutputStreamWriter(outputStream, Charset.forName("UTF-8")));
            //writer.println(xmlTemplate);
            //writer.close();
            outputStream.write(myWord.getBytes());
            outputStream.flush();
            outputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
