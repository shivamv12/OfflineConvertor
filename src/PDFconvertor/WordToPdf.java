/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package PDFconvertor;

import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 *
 * @author milind
 */
public class WordToPdf {

    public static String readDocFile(String fileName) {
        String output = "";
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            HWPFDocument doc = new HWPFDocument(fis);
            WordExtractor we = new WordExtractor(doc);
            String[] paragraphs = we.getParagraphText();
            for (String para : paragraphs) {
                output = output + "\n" + para.toString() + "\n";
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output;
    }

    public static String readDocxFile(String fileName) {
        String output = "";
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph para : paragraphs) {
                output = output + "\n" + para.getText() + "\n";
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output;
    }

    public static void writePdfFile(String output) throws FileNotFoundException, DocumentException {
        File file = new File("D:\\itext-test.pdf");
        FileOutputStream fileout = new FileOutputStream(file);
        Document document = new Document();
        PdfWriter.getInstance(document, fileout);
        document.addAuthor("Milind");
        document.addTitle("My Converted PDF");
        document.open();
        String[] splitter = output.split("\\n");
        for (int i = 0; i < splitter.length; i++) {
            Chunk chunk = new Chunk(splitter[i]);
            Font font = new Font();
            font.setStyle(Font.UNDERLINE);
            font.setStyle(Font.ITALIC);
            chunk.setFont(font);
            document.add(chunk);
            Paragraph paragraph = new Paragraph();
            paragraph.add("");
            document.add(paragraph);
        }
        document.close();

    }

    public static void main(String[] args) throws FileNotFoundException, DocumentException {
        String ext = FilenameUtils.getExtension("D:\\test.docx");
        String output = "";
        if ("docx".equalsIgnoreCase(ext)) {
            output = readDocxFile("D:\\Test.docx");
        } else if ("doc".equalsIgnoreCase(ext)) {
            output = readDocFile("D:\\Test.doc");
        } else {
            System.out.println("INVALID FILE TYPE. ONLY .doc and .docx are permitted.");
        }
        writePdfFile(output);
    }
}