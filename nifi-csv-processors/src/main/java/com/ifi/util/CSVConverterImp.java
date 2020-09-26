package com.ifi.util;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;

public class CSVConverterImp implements CSVConverter {
    private static final Object EXCEL_STYLE_ESCAPING = ""; //TODO
    private String delimiter = ",";
    private String asNullValue = "";
    private FormulaEvaluator evaluator;
    private DataFormatter formatter;
    private Object formattingConvention; //TODO

    public CSVConverterImp() {
    }

    public CSVConverterImp(String delimiter, String asNullValue) {
        this.delimiter = delimiter;
        this.asNullValue = asNullValue;
    }

    public Workbook createWorkbook(InputStream inputStream) throws IOException {

        Workbook workbook = WorkbookFactory.create(inputStream);
        this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        this.formatter = new DataFormatter();
        return workbook;
    }

    public String toCSVFormat(Sheet sheet) {
        StringBuilder builder = new StringBuilder();
        if (sheet.getPhysicalNumberOfRows() <= 0) {
            return asNullValue;
        }
        int lastRowNum = sheet.getLastRowNum();
        for (int j = 0; j <= lastRowNum; j++) {
            Row row = sheet.getRow(j);
            String csvLine = this.rowToCSVFormat(row);
            builder.append(csvLine);
        }
        return builder.toString();
    }

    private String rowToCSVFormat(Row row) {
        StringBuilder builder = new StringBuilder();
        Cell cell;
        int lastCellNum;
        if (row == null) {
            return asNullValue;
        }
        lastCellNum = row.getLastCellNum();
        for (int i = 0; i <= lastCellNum; i++) {
            cell = row.getCell(i);
            String fieldData;
            if (cell.getCellType() != CellType.FORMULA) {
                fieldData = this.formatter.formatCellValue(cell);
            } else {
                fieldData = this.formatter.formatCellValue(cell, this.evaluator);
            }
            builder.append(escapeEmbeddedCharacters(fieldData));
            builder.append(delimiter);
        }
        return builder.toString();
    }

    private String escapeEmbeddedCharacters(String fieldData) {
        StringBuilder builder;
        if (this.formattingConvention == CSVConverterImp.EXCEL_STYLE_ESCAPING) {
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
            return (builder.toString().trim());
        } else {
            if (fieldData.contains(this.delimiter)) {
                fieldData = fieldData.replaceAll(this.delimiter, ("\\\\" + this.delimiter));
            }
            if (fieldData.contains("\n")) {
                fieldData = fieldData.replaceAll("\n", "\\\\\n");
            }
            return (fieldData);
        }
    }
}
