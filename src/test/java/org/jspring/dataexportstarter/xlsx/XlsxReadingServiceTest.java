package org.jspring.dataexportstarter.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.dataexportstarter.xlsx.domain.SheetInfo;
import org.jspring.dataexportstarter.xlsx.service.XlsxReadingService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.Optional;

import static org.jspring.dataexportstarter.xlsx.domain.CellCoordinates.SearchBuilder;

@SpringBootTest
class XlsxReadingServiceTest {

    @Autowired
    private XlsxReadingService xlsxReadingService;

    private SheetInfo sheetInfo;


    @BeforeEach
    public void setUp() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        sheetInfo = new SheetInfo(workbook, "Risconti");
    }

    @Test
    @DisplayName("Search cell by value")
    void testXlsxTargetCell() {

        Optional<Cell> cell = xlsxReadingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("Voce :")
                        .build()
        );

        if (cell.isPresent()) {
            System.out.println("found!!!");
        }

    }

    @Test
    @DisplayName("Search cell by value in first column")
    void testXlsxTargetCellFirstCol() {

        Optional<Cell> cellY = xlsxReadingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue(202311.0)
                        .firstColumn()
                        .build()
        );

        if (cellY.isPresent()) {
            System.out.println("found");
        }

    }


    @Test
    @DisplayName("Search two cells by value")
    void testXlsxTargetCellAndWriteWithNamedCoordinates() {

        Optional<Cell> cellX = xlsxReadingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("Quota")
                        .build()
        );

        Optional<Cell> cellY = xlsxReadingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue(202311)
                        .firstColumn()
                        .build()
        );

        if (cellX.isPresent() && cellY.isPresent()) {
            System.out.println("found coordinates: x -> " + cellX.get().getColumnIndex() + " y -> " + cellY.get().getRowIndex());
        }

    }

    @Test
    @DisplayName("Get cell value by coordinates")
    void testXlsxTargetCellAndWriteWithCoordinates() {

        xlsxReadingService.searchCellBySheetAndCoordinates(
                        sheetInfo,
                        SearchBuilder.init()
                                .rowNumber(2)
                                .columnNumber(0)
                                .build()
                )
                .ifPresent(
                        cell -> System.out.println("found coordinates value: " + xlsxReadingService.readCellValue(cell).value())
                );

    }

    @Test
    @DisplayName("Get cell value by cell address")
    void testXlsxTargetCellAndWriteWithAddress() {

        xlsxReadingService.searchCellBySheetAndCoordinates(
                        sheetInfo,
                        SearchBuilder.init()
                                .address("A2")
                                .build()
                )
                .ifPresent(
                        cell -> System.out.println("found coordinates address value: " + xlsxReadingService.readCellValue(cell).value())
                );

    }


}