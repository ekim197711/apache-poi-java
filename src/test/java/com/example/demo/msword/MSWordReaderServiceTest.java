package com.example.demo.msword;

import org.junit.jupiter.api.Test;

import java.io.IOException;

class MSWordReaderServiceTest {

    @Test
    void readStuff() throws IOException {
        new MSWordReaderService().readStuff();
    }

    @Test
    void printParts() throws IOException {
        new MSWordReaderService().printParts();
    }

    @Test
    void printFooter() throws IOException {
        new MSWordReaderService().printFooter();
    }

    @Test
    void printHeader() throws IOException {
        new MSWordReaderService().printHeader();
    }
}