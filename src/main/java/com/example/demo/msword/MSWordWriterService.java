package com.example.demo.msword;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;

@Service
@Slf4j
public class MSWordWriterService {


    private void saveDocument(XWPFDocument doc) throws IOException {
        File file = new File("/home/mike/IdeaProjects/influencer/apache-poi-excel/input/Generated Information.docx");
        file.delete();
        file.createNewFile();
        doc.write(new FileOutputStream(file));
//        ByteArrayOutputStream baos = new ByteArrayOutputStream(4096);
//        doc.write(baos);
        doc.close();
    }

    public void writeASmallStory() throws IOException {
        XWPFDocument document = new XWPFDocument();
        XWPFDocument template = new XWPFDocument(new FileInputStream("/home/mike/IdeaProjects/influencer/apache-poi-excel/input/Important Information.docx"));
        XWPFStyles newStyles = document.createStyles();
        XWPFStyle heading1 = template.getStyles().getStyle("Heading1");
        XWPFStyle heading2 = template.getStyles().getStyle("Heading2");
        newStyles.addStyle(heading1);
        newStyles.addStyle(heading2);

        generateHeaderAndFooter(document);
        generateBodyParagraphs(document, heading1, heading2);
        generateTable(document);

        saveDocument(document);
    }

    public void generateTable(XWPFDocument document) {
        XWPFTable table = document.createTable(10, 2);
        int rownumber = 0;
        table.getRow(rownumber).getCell(0).setText("Name");
        table.getRow(rownumber++).getCell(1).setText("Age");
        table.getRow(rownumber).getCell(0).setText("Mike");
        table.getRow(rownumber++).getCell(1).setText("42");
        table.getRow(rownumber).getCell(0).setText("Susan");
        table.getRow(rownumber++).getCell(1).setText("26");
        table.getRow(rownumber).getCell(0).setText("John");
        table.getRow(rownumber++).getCell(1).setText("31");
    }

    public void generateHeaderAndFooter(XWPFDocument document) {
        XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);
        header.createParagraph().createRun().setText("Header tralallala 1 2 3 4 5 6");
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph p = footer.createParagraph();
        XWPFRun run = p.createRun();
        run.setFontSize(40);
        run.setText("Footer text bla bla bla");
    }

    public void generateBodyParagraphs(XWPFDocument document, XWPFStyle heading1, XWPFStyle heading2) throws IOException {
        XWPFParagraph pBody = document.createParagraph();
        XWPFRun run1 = pBody.createRun();
        pBody.setStyle(heading1.getStyleId());
        run1.setStyle(heading1.getStyleId());
        run1.setText("This is the first headerr");

        XWPFParagraph pBody2 = document.createParagraph();
        XWPFRun run2 = pBody2.createRun();
        pBody2.setStyle(heading2.getStyleId());
        run2.setText("This is the second headerr");
        run2.addCarriageReturn();
        // Empty paragraph
        document.createParagraph();

    }

}
