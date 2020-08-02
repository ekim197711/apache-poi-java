package com.example.demo;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;


public class CreateSpreadSheetServiceTest {


    @Test
    public void testPlan() throws IOException {
        new CreateSpreadSheetService().writeToExcel();
    }
}