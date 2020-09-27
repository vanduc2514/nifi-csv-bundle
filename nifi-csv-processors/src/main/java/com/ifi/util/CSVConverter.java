package com.ifi.util;

import com.ifi.util.exception.InvalidDocumentException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;

public interface CSVConverter {

    Workbook createWorkbook(InputStream inputStream) throws IOException, InvalidDocumentException, EncryptedDocumentException;

    String toCSVFormat(Sheet sheet);

}
