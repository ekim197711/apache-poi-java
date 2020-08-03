package com.example.demo;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class ReadSpreadSheetServiceTest {
    @Test
    public void testPlan() throws IOException {
        List<ProductSale> productSales = new ReadSpreadSheetService().readWorkBook();
        productSales.forEach(System.out::println);
    }
}