package org.jspring.dataexportstarter.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.dataexportstarter.xlsx.domain.SheetInfo;
import org.jspring.dataexportstarter.xlsx.service.XlsxReadingService;
import org.jspring.dataexportstarter.xlsx.service.XlsxWritingService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.Optional;

import static org.jspring.dataexportstarter.xlsx.domain.CellCoordinates.SearchBuilder;

@SpringBootTest
class XlsxWritingServiceTest {

    @Autowired
    private XlsxReadingService xlsxReadingService;
    @Autowired
    private XlsxWritingService xlsxWritingService;

    private XSSFWorkbook workbook;
    private SheetInfo sheetInfo;


    @BeforeEach
    public void setUp() {
        workbook = xlsxReadingService.readFromTemplate();
        sheetInfo = new SheetInfo(workbook, "Risconti");
    }


    @Test
    @DisplayName("Write to the right of cell searched by value")
    void testXlsxTargetCellAndWriteToTheRight() {

        Optional<Cell> cell = xlsxReadingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("Voce :")
                        .build()
        );

        if (cell.isPresent()) {
            System.out.println("found!!!");
            xlsxWritingService.writeValue(
                    xlsxReadingService.getRightCell(cell.get()),
                    "voce1234",
                    CellType.STRING
            );

            xlsxWritingService.writeFile(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");
        }

    }


    @Test
    @DisplayName("Write at coordinates of two cells searched by value")
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
            xlsxWritingService.writeValue(
                    xlsxReadingService.getCellByCoordinates(cellX.get(), cellY.get()),
                    12.34,
                    CellType.NUMERIC
            );

            xlsxWritingService.writeFile(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");
        }

    }


    @Test
    @DisplayName("Evaluate formula to the right of cell searched by value")
    void testXlsxEvaluateFormulas() {

        Optional<Cell> cell = xlsxReadingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("TOTALI ")
                        .firstColumn()
                        .build()
        );

        if (cell.isPresent()) {
            Cell cellRight = cell.get().getRow().getCell(cell.get().getColumnIndex() + 1);
            cellRight.setCellFormula("SUM(B6:B17)");
            XSSFFormulaEvaluator formulaEvaluator =
                    workbook.getCreationHelper().createFormulaEvaluator();
            formulaEvaluator.evaluate(cellRight);
        }

        xlsxWritingService.writeFile(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");

    }


}