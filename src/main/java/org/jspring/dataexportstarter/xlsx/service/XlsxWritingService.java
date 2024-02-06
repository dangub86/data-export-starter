package org.jspring.dataexportstarter.xlsx.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class XlsxWritingService {

    public void writeFile(XSSFWorkbook workbook, String fileName) {
        try (FileOutputStream out = new FileOutputStream(fileName)) {
            workbook.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public byte[] writeAsByteArray(XSSFWorkbook workbook) {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            workbook.write(baos);
            return baos.toByteArray();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void writeValue(
            Cell cell, Object value) {

        switch (value) {
            case String stringVal -> cell.setCellValue(stringVal);
            case Double doubleVal -> cell.setCellValue(doubleVal);
            case Boolean booleanVal -> cell.setCellValue(booleanVal);
            default -> throw new IllegalStateException("Unexpected value type: " + value.getClass().getName());
        }

    }


}
