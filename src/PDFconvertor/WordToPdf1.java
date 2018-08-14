package PDFconvertor;
/** * * @author shibabrata.b */
import java.io.*;
import java.awt.*;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import java.io.*;
import java.util.zip.*;
import javax.swing.text.Document;
public class WordToPdf1 {
    public static void main(String arg[])throws Exception {
        System.out.println(""Hello RoseIndia"");
        Document document = new Document(PageSize.A4, 36, 72, 108, 180);
        PdfWriter.getInstance(document,System.out);
        PdfWriter.getInstance(document,new FileOutputStream(""C:/shib/PHP SQL UNIX INTERVIEW HELPER.pdf""));
        document.open();
        ZipInputStream zip = new ZipInputStream(new BufferedInputStream(new FileInputStream(""C:/shib/PHP SQL UNIX INTERVIEW HELPER.zip"")));
        ZipEntry entry;
        while((entry = zip.getNextEntry()) != null) {
            byte data[] = new byte[1024];
            int count;
            String text="""";
            while((count=zip.read(data,0,1024))!=-1) {
                text=new String(data);
                document.add(new Paragraph(text));
                }
            }
        document.close(); } } 