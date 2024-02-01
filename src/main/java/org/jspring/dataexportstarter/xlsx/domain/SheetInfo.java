package org.jspring.dataexportstarter.xlsx.domain;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public record SheetInfo(
        XSSFWorkbook workbook,
        String sheetName
) {

    public XSSFSheet getSheet() {
        return workbook().getSheet(sheetName());
    }

}
