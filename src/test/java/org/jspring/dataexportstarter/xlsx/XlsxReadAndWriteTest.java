package org.jspring.dataexportstarter.xlsx;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.dataexportstarter.xlsx.domain.SheetInfo;
import org.jspring.dataexportstarter.xlsx.service.XlsxReadingService;
import org.jspring.dataexportstarter.xlsx.service.XlsxWritingService;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import static org.junit.jupiter.api.Assertions.assertNotNull;

@SpringBootTest
class XlsxReadAndWriteTest {

    @Autowired
    private XlsxReadingService xlsxReadingService;
    @Autowired
    private XlsxWritingService xlsxWritingService;


    @Test
    @DisplayName("Read Xlsx from template")
    void readFromTemplate() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, "Risconti");
        assertNotNull(sheetInfo);
    }

    @Test
    @DisplayName("Write Xlsx to file")
    void writeToFile() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, "Risconti");
        assertNotNull(sheetInfo);
        xlsxWritingService.write(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");

        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate("src/main/resources/xlsx/output/risconti-mod.xls");
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, "Risconti");
        assertNotNull(sheetInfoMod);
    }

    @Test
    @DisplayName("Write and read Xlsx from byte array")
    void writeAndReadFromByteArray() {
        XSSFWorkbook workbookFromTemplate = xlsxReadingService.readFromTemplate();
        byte[] bytes = xlsxWritingService.writeAsByteArray(workbookFromTemplate);
        assertNotNull(bytes);

        XSSFWorkbook workbookFromByteArray = xlsxReadingService.readFromByteArray(bytes);
        SheetInfo sheetInfoFromByteArray = new SheetInfo(workbookFromByteArray, "Risconti");
        assertNotNull(sheetInfoFromByteArray);
    }




}