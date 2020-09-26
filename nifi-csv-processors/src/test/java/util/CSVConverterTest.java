package util;

import com.ifi.util.CSVConverter;
import com.ifi.util.CSVConverterImp;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Objects;

import static org.junit.Assert.*;

public class CSVConverterTest {
    private CSVConverter converter;

    @Before
    public void setUp() {
        converter = new CSVConverterImp();
    }

    @Test
    public void should_read_workbook_as_hssfWorkbook() {
        int expectedSheet = 1;
        int firstSheetIndex = 0;
        int expectedRowInFirstSheet = 501;
        String fileName = "one-sheet-no-formula-972003.xls";
        File file = new File(Objects.requireNonNull(getClass().getClassLoader().getResource(fileName)).getFile());
        assertNotNull(file);
        Workbook workbook = null;
        try {
            workbook = converter.createWorkbook(new FileInputStream(file));
            assertTrue(workbook instanceof HSSFWorkbook);
            assertEquals(expectedSheet, workbook.getNumberOfSheets());
            assertEquals(expectedRowInFirstSheet, workbook.getSheetAt(firstSheetIndex).getPhysicalNumberOfRows());
        } catch (IOException exception) {
            exception.printStackTrace();
        }
        assertNotNull(workbook);
    }

    @Test
    public void should_read_workbook_as_xssfWorkbook() {
        int expectedSheet = 1;
        int firstSheetIndex = 0;
        int expectedRowInFirstSheet = 501;
        String fileName = "one-sheet-no-formula-972003.xls";
        File file = new File(Objects.requireNonNull(getClass().getClassLoader().getResource(fileName)).getFile());
        assertNotNull(file);
        Workbook workbook = null;
        try {
            workbook = converter.createWorkbook(new FileInputStream(file));
            assertTrue(workbook instanceof HSSFWorkbook);
            assertEquals(expectedSheet, workbook.getNumberOfSheets());
            assertEquals(expectedRowInFirstSheet, workbook.getSheetAt(firstSheetIndex).getPhysicalNumberOfRows());
        } catch (IOException exception) {
            exception.printStackTrace();
        }
        assertNotNull(workbook);
    }
}
