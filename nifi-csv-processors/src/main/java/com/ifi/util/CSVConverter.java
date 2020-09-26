package com.ifi.util;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;

public interface CSVConverter {

    Workbook createWorkbook(InputStream inputStream) throws IOException;

    String toCSVFormat(Sheet sheet);

}
