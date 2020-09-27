package util;

import com.ifi.util.CSVConverter;
import com.ifi.util.CSVConverterImp;
import com.ifi.util.exception.InvalidDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Objects;

import static org.junit.Assert.*;

public class CSVConverterTest {
    private CSVConverter converter;
    private static File hssfWorkBookFile;
    private static File xssfWorkBookFile;

    static {
        hssfWorkBookFile = new File(Objects.requireNonNull(CSVConverterTest.class.getClassLoader().getResource(
                "one-sheet-no-formula-972003.xls")).getFile());
        xssfWorkBookFile = new File(Objects.requireNonNull(CSVConverterTest.class.getClassLoader().getResource(
                "one-sheet-no-formula-2007.xlsx")).getFile());

    }

    @Before
    public void setUp() {
        converter = new CSVConverterImp();
    }

    @Test
    public void should_read_workbook_as_hssfWorkbook() {
        int expectedSheet = 1;
        int firstSheetIndex = 0;
        int expectedRowInFirstSheet = 501;

        Workbook workbook = null;
        try {
            workbook = converter.createWorkbook(new FileInputStream(hssfWorkBookFile));
            assertTrue(workbook instanceof HSSFWorkbook);
            assertEquals(expectedSheet, workbook.getNumberOfSheets());
            assertEquals(expectedRowInFirstSheet, workbook.getSheetAt(firstSheetIndex).getPhysicalNumberOfRows());
        } catch (IOException | InvalidDocumentException exception) {
            exception.printStackTrace();
        }
        assertNotNull(workbook);
    }

    @Test
    public void should_read_workbook_as_xssfWorkbook() {
        int expectedSheet = 1;
        int firstSheetIndex = 0;
        int expectedRowInFirstSheet = 10;

        Workbook workbook = null;
        try {
            workbook = converter.createWorkbook(new FileInputStream(xssfWorkBookFile));

            assertTrue(workbook instanceof XSSFWorkbook);
            assertEquals(expectedSheet, workbook.getNumberOfSheets());
            assertEquals(expectedRowInFirstSheet, workbook.getSheetAt(firstSheetIndex).getPhysicalNumberOfRows());
        } catch (IOException | InvalidDocumentException exception) {
            exception.printStackTrace();
        }
        assertNotNull(workbook);
    }

    @Test
    public void should_not_read_as_work_book() {
        String fileName = "not-a-workbook.xls";
        File file = new File(Objects.requireNonNull(getClass().getClassLoader().getResource(fileName)).getFile());

        assertThrows(InvalidDocumentException.class, () -> converter.createWorkbook(new FileInputStream(file)));
    }

    @Test
    public void should_convert_xssf_to_csv() {

    }
}
