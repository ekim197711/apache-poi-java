package com.example.demo;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

public class ReadSpreadSheetServiceTest {
    @Test
    public void testPlan() throws IOException {
        new ReadSpreadSheetService().readWorkBook();
    }
}