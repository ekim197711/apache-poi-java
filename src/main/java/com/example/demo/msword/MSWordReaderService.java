package com.example.demo.msword;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

@Service
@Slf4j
public class MSWordReaderService {


    private XWPFDocument getDocument() throws IOException {
        XWPFDocument doc = new XWPFDocument(
                new FileInputStream(
                        new File("/home/mike/IdeaProjects/influencer/apache-poi-excel/input/Important Information.docx")));
        return doc;
    }

    public void readStuff() throws IOException {
        printHeader();
        printFooter();
        printParts();
    }

    public void printParts() throws IOException {
        XWPFDocument doc = getDocument();
        List<IBodyElement> bodyElements = doc.getBodyElements();
        for (IBodyElement bodyElement : bodyElements) {
            log.info("Body part type: {}", bodyElement.getClass());
            if ( bodyElement instanceof XWPFParagraph)
            {
                XWPFParagraph para = (XWPFParagraph)bodyElement;
                log.info("paragraph text: {}", para.getText());
            }
            else if ( bodyElement instanceof XWPFTable)
            {
                XWPFTable table = (XWPFTable)bodyElement;
                log.info("table text: {}", table.getText());
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows) {
                    List<XWPFTableCell> tableCells = row.getTableCells();
                    log.info("new row with cells: {}",  row.getTableCells().size());
                    for (XWPFTableCell tableCell : tableCells) {
                        log.info("recurs text cell: {} ", tableCell.getText());
                    }
                }
            }
        }
    }

    public void printFooter() throws IOException {
        XWPFDocument doc = getDocument();
        List<XWPFFooter> footerList = doc.getFooterList();
        for (XWPFFooter footer : footerList) {
            log.info("Footer: {}", footer.getText());
        }
    }

    public void printHeader() throws IOException {
        XWPFDocument doc = getDocument();
        List<XWPFHeader> headerList = doc.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            log.info("Header: {}", xwpfHeader.getText());
        }
    }


}
