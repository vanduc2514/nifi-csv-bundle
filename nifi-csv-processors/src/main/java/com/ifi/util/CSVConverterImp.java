package com.ifi.util;

import com.ifi.util.exception.InvalidDocumentException;
import org.apache.poi.EmptyFileException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;

public class CSVConverterImp implements CSVConverter {
    private String delimiter = ",";
    private final String EMPTY_STRING = "";
    private FormulaEvaluator evaluator;
    private DataFormatter formatter;
    private EscapeChar escapeChar = EscapeChar.EXCEL_STYLE_ESCAPING;

    public CSVConverterImp() {
    }

    public CSVConverterImp(String delimiter, EscapeChar escapeChar) {
        this.delimiter = delimiter;
        this.escapeChar = escapeChar;
    }

    public Workbook createWorkbook(InputStream inputStream) throws IOException, InvalidDocumentException, EncryptedDocumentException {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(inputStream);
            this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            this.formatter = new DataFormatter();
        } catch (EmptyFileException exception) {
            throw new InvalidDocumentException();
        }
        return workbook;
    }

    public String toCSVFormat(Sheet sheet) {
        StringBuilder builder = new StringBuilder();
        if (sheet.getPhysicalNumberOfRows() <= 0) {
            return EMPTY_STRING;
        }
        int lastRowNum = sheet.getLastRowNum();
        for (int j = 0; j <= lastRowNum; j++) {
            Row row = sheet.getRow(j);
            String csvLine = this.rowToCSVFormat(row);
            builder.append(csvLine);
            builder.append("\n");
        }
        return builder.toString();
    }

    private String rowToCSVFormat(Row row) {
        StringBuilder builder = new StringBuilder();
        Cell cell;
        int lastCellNum;
        if (row == null) {
            return EMPTY_STRING;
        }
        lastCellNum = row.getLastCellNum();
        for (int i = 0; i < lastCellNum; i++) {
            cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            String fieldData;
            if (cell.getCellType() != CellType.FORMULA) {
                fieldData = this.formatter.formatCellValue(cell);
            } else {
                fieldData = this.formatter.formatCellValue(cell, this.evaluator);
            }
            builder.append(escapeEmbeddedCharacters(fieldData));
            if (i != lastCellNum - 1) {
                builder.append(delimiter);
            }
        }
        return builder.toString();
    }

    private String escapeEmbeddedCharacters(String fieldData) {
        StringBuilder builder;
        if (this.escapeChar == EscapeChar.EXCEL_STYLE_ESCAPING) {
            if (fieldData.contains("\"")) {
                builder = new StringBuilder(fieldData.replaceAll("\"", "\\\"\\\""));
                builder.insert(0, "\"");
                builder.append("\"");
            } else {
                builder = new StringBuilder(fieldData);
                if ((builder.indexOf(this.delimiter)) > -1 ||
                        (builder.indexOf("\n")) > -1) {
                    builder.insert(0, "\"");
                    builder.append("\"");
                }
            }
            return builder.toString();
        } else if (escapeChar == EscapeChar.UNIX_STYLE_ESCAPING) {
            if (fieldData.contains(this.delimiter)) {
                fieldData = fieldData.replaceAll(this.delimiter, ("\\\\" + this.delimiter));
            }
            if (fieldData.contains("\n")) {
                fieldData = fieldData.replaceAll("\n", "\\\\\n");
            }
            return (fieldData);
        } else {
            return (fieldData);
        }
    }
}
