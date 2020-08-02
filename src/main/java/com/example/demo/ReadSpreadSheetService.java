package com.example.demo;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;

@Service
@Slf4j
public class ReadSpreadSheetService {

    public String readWorkBook() throws IOException {
        Workbook wb = new XSSFWorkbook(new FileInputStream("input/business_input_data.ods"));
        Sheet sheet = wb.getSheet("Importantdata");
        Row headlines = sheet.getRow(1);
        String header1 = headlines.getCell(0).getStringCellValue();
        String header2 = headlines.getCell(0).getStringCellValue();
        log.info("Header 1 {} Header 2 {}", header1, header2);
        return "done";
    }


}
