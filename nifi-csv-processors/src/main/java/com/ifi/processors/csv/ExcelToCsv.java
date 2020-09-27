/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.ifi.processors.csv;

import com.ifi.util.CSVConverter;
import com.ifi.util.CSVConverterImp;
import com.ifi.util.EscapeChar;
import com.ifi.util.exception.InvalidDocumentException;
import org.apache.commons.io.FilenameUtils;
import org.apache.nifi.annotation.behavior.ReadsAttribute;
import org.apache.nifi.annotation.behavior.ReadsAttributes;
import org.apache.nifi.annotation.behavior.WritesAttribute;
import org.apache.nifi.annotation.behavior.WritesAttributes;
import org.apache.nifi.annotation.documentation.CapabilityDescription;
import org.apache.nifi.annotation.documentation.SeeAlso;
import org.apache.nifi.annotation.documentation.Tags;
import org.apache.nifi.annotation.lifecycle.OnScheduled;
import org.apache.nifi.components.PropertyDescriptor;
import org.apache.nifi.flowfile.FlowFile;
import org.apache.nifi.flowfile.attributes.CoreAttributes;
import org.apache.nifi.logging.ComponentLog;
import org.apache.nifi.processor.*;
import org.apache.nifi.processor.exception.ProcessException;
import org.apache.nifi.util.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.nio.charset.StandardCharsets;
import java.util.*;

@Tags({"excel, csv"})
@CapabilityDescription("Processor to convert Excel to CSV")
@SeeAlso()
@ReadsAttributes({@ReadsAttribute(attribute = "", description = "")})
@WritesAttributes({@WritesAttribute(attribute = "", description = "")})
public class ExcelToCsv extends AbstractProcessor {
    private List<PropertyDescriptor> descriptors;
    private Set<Relationship> relationships;

    static final String UNIX_SYSTEM = "Unix";
    static final String WINDOWS_SYSTEM = "Windows";
    static final String CSV_MIME_TYPE = "text/csv";
    static final String SHEET_NAME_SEPARATOR = "-";
    static final String CSV_EXTENSION = ".csv";
    static final String SHEET_NAME_ATT = "sheet name";
    static final String ROW_NUM_ATT = "row num";
    static final String SOURCE_NAME_ATT = "source name";
    static final byte[] BYTE_ORDER_MARKER = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};

    private ComponentLog logger;
    private CSVConverter converter;
    private boolean utf8Encoded;

    public static final PropertyDescriptor UTF8_ENCODED = new PropertyDescriptor
            .Builder().name("utf8-encoded")
            .displayName("UTF-8 Encode")
            .description("Should the csv file be encoded in UTF-8 charset")
            .allowableValues("true", "false")
            .defaultValue("false")
            .required(true)
            .build();

    public static final PropertyDescriptor ESCAPE_CONVENTION = new PropertyDescriptor
            .Builder().name("escape-convention")
            .displayName("Escape Convention")
            .description("Choose operating system that special character should be escaped")
            .allowableValues(UNIX_SYSTEM, WINDOWS_SYSTEM)
            .defaultValue("Windows")
            .required(true)
            .build();

    public static final PropertyDescriptor DELIMITER = new PropertyDescriptor
            .Builder().name("delimiter")
            .displayName("Delimiter")
            .description("Delimiter use in csv")
            .defaultValue(",")
            .required(true)
            .build();

    public static final Relationship SUCCESS = new Relationship.Builder()
            .name("success")
            .description("Excel files that have been successfully converted to csv are transferred to this relationship")
            .build();

    public static final Relationship FAILURE = new Relationship.Builder()
            .name("failure")
            .description("Excel files that can't be converted to csv are transferred to this relationship")
            .build();

    public static final Relationship ORIGINAL = new Relationship.Builder()
            .name("original")
            .description("Original Excel Files are transferred to this relationship")
            .build();

    @Override
    protected void init(final ProcessorInitializationContext context) {
        final List<PropertyDescriptor> descriptors = new ArrayList<>();
        descriptors.add(UTF8_ENCODED);
        descriptors.add(ESCAPE_CONVENTION);
        descriptors.add(DELIMITER);
        this.descriptors = Collections.unmodifiableList(descriptors);

        final Set<Relationship> relationships = new HashSet<>();
        relationships.add(SUCCESS);
        relationships.add(FAILURE);
        relationships.add(ORIGINAL);
        this.relationships = Collections.unmodifiableSet(relationships);
    }

    @Override
    public Set<Relationship> getRelationships() {
        return this.relationships;
    }

    @Override
    public final List<PropertyDescriptor> getSupportedPropertyDescriptors() {
        return descriptors;
    }

    @OnScheduled
    public void onScheduled(final ProcessContext context) {
        logger = getLogger();
        logger.info("Processor Started!");
        utf8Encoded = context.getProperty(UTF8_ENCODED).asBoolean();
        setupConverter(context);
    }

    private void setupConverter(ProcessContext context) {
        EscapeChar escapeChar = EscapeChar.EXCEL_STYLE_ESCAPING;
        String convention = context.getProperty(ESCAPE_CONVENTION).getValue();
        if (convention.equals(UNIX_SYSTEM)) {
            escapeChar = EscapeChar.UNIX_STYLE_ESCAPING;
        }
        String delimiter = context.getProperty(DELIMITER).getValue();
        converter = new CSVConverterImp(delimiter, escapeChar);
    }

    @Override
    public void onTrigger(final ProcessContext context, final ProcessSession session) throws ProcessException {
        logger.info("Processor Triggered!");
        FlowFile excelFile = session.get();
        if (excelFile == null) {
            return;
        }

        session.read(excelFile, inputStream -> {
            try {
                Workbook workbook = converter.createWorkbook(inputStream);
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String csvData = converter.toCSVFormat(sheet);
                    handleCSVData(session, excelFile, sheet, csvData);
                }
            } catch (InvalidDocumentException e) {
                session.transfer(excelFile, FAILURE);
            }
        });

        session.transfer(excelFile, ORIGINAL);
    }

    private void handleCSVData(ProcessSession session, FlowFile excelFile, Sheet sheet, String csvData) {
        FlowFile csvFile = session.create(excelFile);
        csvFile = session.write(csvFile, outputStream -> {
            if (utf8Encoded) {
                outputStream.write(BYTE_ORDER_MARKER);
                outputStream.write(csvData.getBytes(StandardCharsets.UTF_8));
            } else {
                outputStream.write(csvData.getBytes());
            }
        });

        String sourceFileName = excelFile.getAttribute(CoreAttributes.FILENAME.key());
        csvFile = session.putAttribute(csvFile, SHEET_NAME_ATT, sheet.getSheetName());
        csvFile = session.putAttribute(csvFile, ROW_NUM_ATT, String.valueOf(sheet.getPhysicalNumberOfRows()));
        csvFile = session.putAttribute(csvFile, SOURCE_NAME_ATT, sourceFileName);
        csvFile = session.putAttribute(csvFile, CoreAttributes.MIME_TYPE.key(), CSV_MIME_TYPE);
        csvFile = session.putAttribute(csvFile, CoreAttributes.FILENAME.key(),
                StringUtils.isNotEmpty(sourceFileName) ?
                        getCSVFileName(sourceFileName, sheet.getSheetName()) :
                        csvFile.getAttribute(CoreAttributes.UUID.key()) + CSV_EXTENSION);
        session.transfer(csvFile, SUCCESS);
    }

    private String getCSVFileName(String sourceFileName, String sheetName) {
        StringBuilder builder = new StringBuilder();
        String ext = FilenameUtils.getExtension(sourceFileName);
        if (StringUtils.isNotEmpty(ext)) {
            builder.append(sourceFileName.replaceAll(("." + ext), ""));
        } else {
            builder.append(sourceFileName);
        }
        builder.append(SHEET_NAME_SEPARATOR);
        builder.append(sheetName);
        builder.append(CSV_EXTENSION);
        return builder.toString();
    }
}
