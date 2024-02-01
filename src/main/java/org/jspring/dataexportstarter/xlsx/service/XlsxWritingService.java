package org.jspring.dataexportstarter.xlsx.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class XlsxWritingService {

    public void write(XSSFWorkbook workbook, String fileName) {
        try (FileOutputStream out = new FileOutputStream(fileName)) {
            workbook.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void writeToTheRightOfTheCell(
            Cell cell,
            String value
    ) {
        Cell cell1 = cell.getRow().getCell(cell.getColumnIndex() + 1);
        cell1.setCellValue(value);


    }

    public void writeWithCellCoordinates(
            Cell cellX, Cell cellY, String value) {

        Cell cell = cellY.getRow().getCell(
                cellX.getColumnIndex()
        );

        cell.setCellValue(value);

    }

    public void writeWithCellCoordinates(
            Cell cellX, Cell cellY, double value) {

        Cell cell = cellY.getRow().getCell(
                cellX.getColumnIndex()
        );

        cell.setCellValue(value);

    }


}