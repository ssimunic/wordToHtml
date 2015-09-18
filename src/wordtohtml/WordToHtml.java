package wordtohtml;

import java.io.BufferedWriter;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.StringWriter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.JEditorPane;
import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.w3c.dom.Document;

public class WordToHtml {

    private File docFile;
    private File file;

    public WordToHtml(File docFile) {
        this.docFile = docFile;
    }

    public void convert(File file) {
        this.file = file;

        try {
            FileInputStream finStream = new FileInputStream(docFile.getAbsolutePath());
            HWPFDocument doc = new HWPFDocument(finStream);
            WordExtractor wordExtract = new WordExtractor(doc);
            Document newDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(newDocument);
            wordToHtmlConverter.processDocument(doc);

            StringWriter stringWriter = new StringWriter();
            Transformer transformer = TransformerFactory.newInstance().newTransformer();

            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            transformer.setOutputProperty(OutputKeys.METHOD, "html");
            transformer.transform(new DOMSource(wordToHtmlConverter.getDocument()), new StreamResult(stringWriter));

            String html = stringWriter.toString();
            String newFileName = file.toString().replace(".doc", ".html");
            
            new File("html").mkdirs();
  
            FileOutputStream fos = new FileOutputStream(new File("html/"+newFileName));
            DataOutputStream dos;

            try {
                BufferedWriter out = new BufferedWriter(new OutputStreamWriter(fos, "UTF-8"));
                out.write(html);
                out.close();
            } catch (IOException e) {
                System.out.println("0");
            }

        } catch (Exception e) {
            System.out.println("0");
        }
    }

    public static void main(String args[]) {
        try {
            WordToHtml wordtohtml = new WordToHtml(new File(args[0]));
            wordtohtml.convert(wordtohtml.docFile);
            
        } catch (Exception e) {
            System.out.println("0");
        } finally {
            System.out.println("1");
        }

       
    }
}
