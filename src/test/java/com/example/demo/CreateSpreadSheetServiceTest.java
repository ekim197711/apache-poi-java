package com.example.demo;

import org.junit.jupiter.api.Test;

import java.io.IOException;


public class CreateSpreadSheetServiceTest {


    @Test
    public void testPlan() throws IOException {
        new CreateSpreadSheetService().writeToExcel();
    }
}