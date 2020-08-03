package com.example.demo;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
@Slf4j
public class ReadSpreadSheetService {

    public List<ProductSale> readWorkBook() throws IOException {
        List<ProductSale> sales = new ArrayList<>();

        Workbook wb = new XSSFWorkbook(new FileInputStream("input/business_input_data.xlsx"));
        Sheet sheet = wb.getSheet("Importantdata");
        Row headlines = sheet.getRow(2);
        String header1 = headlines.getCell(0).getStringCellValue();
        String header2 = headlines.getCell(1).getStringCellValue();
        log.info("Header 1 {} Header 2 {}", header1, header2);

        int rowno = 3;
        String product = "NOT_SET";
        Double value = 0d;

        while (product != null && product.length() > 0) {
            Row row = sheet.getRow(rowno++);
            if (row == null)
                break;
            product = row.getCell(0).getStringCellValue();
            value = 0d;
            Cell cell = row.getCell(1);
//            log.info("product {}, value {}", product, cell.toString());
            switch (cell.getCellType()) {
                case _NONE:
                    break;
                case NUMERIC:
                    if (rowno != 6)
                        value = cell.getNumericCellValue();
                    else
                        log.info("We have a date!.... {}", cell.getDateCellValue());
                    break;
                case STRING:
                    value = 0.0d + cell.getStringCellValue().length();
                    break;
                case FORMULA:
                    break;
                case BLANK:
                    break;
                case BOOLEAN:
                    value = cell.getBooleanCellValue() ? 100d : 0d;
                    break;
                case ERROR:
                    break;
            }
            sales.add(new ProductSale(product, value));
        }
        return sales;
    }


}
