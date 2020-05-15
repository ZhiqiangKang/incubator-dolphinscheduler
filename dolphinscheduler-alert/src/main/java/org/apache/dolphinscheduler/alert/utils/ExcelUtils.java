/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.apache.dolphinscheduler.alert.utils;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SequenceWriter;
import com.fasterxml.jackson.dataformat.csv.CsvMapper;
import com.fasterxml.jackson.dataformat.csv.CsvSchema;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * excel utils
 */
public class ExcelUtils {

    private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * generate csv file
     *
     * @param filePath
     * @param content
     * @throws IOException
     */
    public static void genCsvFile(String filePath, String content) throws IOException {

        final JsonNode jsonNode = new ObjectMapper().readTree(content);

        CsvSchema.Builder csvSchemaBuilder = CsvSchema.builder();
        JsonNode firstObject = jsonNode.elements().next();
        firstObject.fieldNames().forEachRemaining(fieldName -> {csvSchemaBuilder.addColumn(fieldName);} );
        CsvSchema csvSchema = csvSchemaBuilder.build().withHeader();

        CsvMapper csvMapper = new CsvMapper();

        FileOutputStream fileOutputStream = null;
        SequenceWriter writer = null;
        try{
            logger.debug("generate csv file: {} begin", filePath);
            fileOutputStream = new FileOutputStream(new File(filePath));
            // 写入文件BOM头
            fileOutputStream.write(new byte[]{(byte)0xEF, (byte)0xBB, (byte)0xBF});
            writer = csvMapper.writerFor(JsonNode.class)
                    .with(csvSchema)
                    .writeValues(fileOutputStream)
                    .write(jsonNode);
        } finally {
            if (Objects.nonNull(writer)){
                writer.flush();
                writer.close();
            }

            if (Objects.nonNull(fileOutputStream)){
                fileOutputStream.flush();
                fileOutputStream.close();
            }
        }

        logger.debug("generate csv file: {} end", filePath);
    }


    /**
     * generate excel file
     *
     * @param filePath
     * @param content
     */
    public static void genExcelFile(String filePath, String content){
        List<LinkedHashMap> itemsList;

        try {
            itemsList = JSONUtils.toList(content, LinkedHashMap.class);
        }catch (Exception e){
            logger.error(String.format("json format incorrect : %s",content),e);
            throw new RuntimeException("json format incorrect",e);
        }

        if (itemsList == null || itemsList.size() == 0){
            logger.error("itemsList is null");
            throw new RuntimeException("itemsList is null");
        }

        logger.info("itemsList.size: {}", itemsList.size());
        Workbook wb = null;
        FileOutputStream fos = null;
        try {
            LinkedHashMap<String, Object> headerMap = itemsList.get(0);

            Iterator<Map.Entry<String, Object>> iter = headerMap.entrySet().iterator();
            List<String> headerList = new ArrayList<>();
            while (iter.hasNext()){
                Map.Entry<String, Object> en = iter.next();
                headerList.add(en.getKey());
                logger.debug("en.getKey(): {}", en.getKey());
            }

            wb = new HSSFWorkbook();
            logger.debug("wb: {}", wb);

            // generate a table
            Sheet sheet = wb.createSheet();
            logger.debug("createSheet");
            Row row = sheet.createRow(0);
            //set the height of the first line
            row.setHeight((short)500);

            //setting excel headers
            logger.debug("setting excel headers begin");
            for (int i = 0; i < headerList.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(headerList.get(i));
            }
            logger.debug("setting excel headers end");

            //setting excel body
            logger.debug("setting excel body begin");
            int rowIndex = 1;
            for (LinkedHashMap<String, Object> itemsMap : itemsList){
                Object[] values = itemsMap.values().toArray();
                row = sheet.createRow(rowIndex);
                logger.debug("createRow, rowIndex: {}", rowIndex);
                //setting excel body height
                row.setHeight((short)500);
                rowIndex++;
                for (int j = 0 ; j < values.length ; j++){
                    Cell cell1 = row.createCell(j);
                    cell1.setCellValue(String.valueOf(values[j]));
                }
            }
            logger.debug("setting excel body end");

            for (int i = 0; i < headerList.size(); i++) {
                sheet.setColumnWidth(i, headerList.get(i).length() * 800);

            }

            //setting file output
            fos = new FileOutputStream(filePath);

            wb.write(fos);

        }catch (Exception e){
            logger.error("generate excel error",e);
            throw new RuntimeException("generate excel error",e);
        }finally {
            if (wb != null){
                try {
                    wb.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
            }
            if (fos != null){
                try {
                    fos.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
            }
        }
    }

}
