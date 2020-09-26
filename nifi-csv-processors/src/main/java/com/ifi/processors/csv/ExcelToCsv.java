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
import org.apache.nifi.logging.ComponentLog;
import org.apache.nifi.processor.*;
import org.apache.nifi.processor.exception.ProcessException;

import java.io.ByteArrayOutputStream;
import java.util.*;

@Tags({"excel, csv"})
@CapabilityDescription("Processor to convert Excel to CSV")
@SeeAlso()
@ReadsAttributes({@ReadsAttribute(attribute = "", description = "")})
@WritesAttributes({@WritesAttribute(attribute = "", description = "")})
public class ExcelToCsv extends AbstractProcessor {
    private ComponentLog logger;
//    public static final PropertyDescriptor PROCESSOR_PROPERTY = new PropertyDescriptor
//            .Builder().name("MY_PROPERTY")
//            .displayName("My property")
//            .description("Example Property")
//            .required(true)
//            .addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
//            .build();

    public static final Relationship SUCCESS = new Relationship.Builder()
            .name("success")
            .description("Excel files that have been successfully converted to CSV are transferred to this relationship")
            .build();

    public static final Relationship FAILURE = new Relationship.Builder()
            .name("failure")
            .description("Excel files that can't be converted to CSV are transferred to this relationship")
            .build();

    private List<PropertyDescriptor> descriptors;

    private Set<Relationship> relationships;

    @Override
    protected void init(final ProcessorInitializationContext context) {
//        final List<PropertyDescriptor> descriptors = new ArrayList<>();
//        descriptors.add(PROCESSOR_PROPERTY);
//        this.descriptors = Collections.unmodifiableList(descriptors);

        final Set<Relationship> relationships = new HashSet<>();
        relationships.add(SUCCESS);
        relationships.add(FAILURE);
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
    }

    @Override
    public void onTrigger(final ProcessContext context, final ProcessSession session) throws ProcessException {
        logger.info("Processor Triggered!");
        FlowFile flowFile = session.get();
        if (flowFile == null) {
            return;
        }
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        session.exportTo(flowFile, outputStream);
        final byte[] bytebuffer = outputStream.toByteArray();
        logger.info(Arrays.toString(bytebuffer));
        session.transfer(flowFile, SUCCESS);
    }
}
