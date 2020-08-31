package com.example.demo.msword;

import org.junit.jupiter.api.Test;
import java.io.IOException;
import static org.junit.jupiter.api.Assertions.*;

class MSWordWriterServiceTest {

    @Test
    void writeASmallStory() throws IOException {
        new MSWordWriterService().writeASmallStory();
    }
}