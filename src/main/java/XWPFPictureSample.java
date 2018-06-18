import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

public class XWPFPictureSample {
    public static void main(String[] args) throws Exception {
        InputStream dis = new FileInputStream("src/main/resources/template.docx");
        XWPFDocument doc = new XWPFDocument(dis);

        XWPFParagraph title = doc.createParagraph();
        XWPFRun run = title.createRun();
        run.setText("Fig.1 Sample");
        run.setBold(true);
        title.setAlignment(ParagraphAlignment.CENTER);

        InputStream iis = new FileInputStream("src/main/resources/sample.png");
        run.addBreak();
        run.addPicture(iis, XWPFDocument.PICTURE_TYPE_PNG, "sample.png", Units.toEMU(300), Units.toEMU(112)); // 300x112 pixels
        iis.close();

        OutputStream dos = new FileOutputStream("sample.docx");
        doc.write(dos);
        dos.close();
    }
}
