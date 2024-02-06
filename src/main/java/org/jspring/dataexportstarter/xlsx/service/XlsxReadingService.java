package org.jspring.dataexportstarter.xlsx.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.dataexportstarter.xlsx.domain.CellCoordinates;
import org.jspring.dataexportstarter.xlsx.domain.CellSearch;
import org.jspring.dataexportstarter.xlsx.domain.CellWrapper;
import org.jspring.dataexportstarter.xlsx.domain.SheetInfo;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Optional;


public class XlsxReadingService {

    private final String templatePath;
    public XlsxReadingService(String templatePath) {
        this.templatePath = templatePath;
    }

    public XSSFWorkbook readFromTemplate() {
        try (FileInputStream inputStream = new FileInputStream(templatePath)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public XSSFWorkbook readFromTemplate(String templatePath) {
        try (FileInputStream inputStream = new FileInputStream(templatePath)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public XSSFWorkbook readFromByteArray(byte[] fileContent) {
        try (ByteArrayInputStream inputStream = new ByteArrayInputStream(fileContent)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public <T> Optional<Cell> search(
            SheetInfo sheetInfo,
            CellCoordinates<T> cellCoordinates
    ) {

        return new CellSearch<>(sheetInfo, cellCoordinates)
                .search();

    }


    public CellWrapper<?> readValue(Cell cell) {

        return switch (cell.getCellType()) {
            case STRING -> new CellWrapper<>(cell.getCellType(), cell.getStringCellValue());
            case NUMERIC -> new CellWrapper<>(cell.getCellType(), cell.getNumericCellValue());
            case BOOLEAN -> new CellWrapper<>(cell.getCellType(), cell.getBooleanCellValue());
            case FORMULA -> new CellWrapper<>(cell.getCellType(), cell.getCellFormula());
            case BLANK, _NONE, ERROR -> new CellWrapper<>(cell.getCellType(), "Unhandled");
        };
    }
}
