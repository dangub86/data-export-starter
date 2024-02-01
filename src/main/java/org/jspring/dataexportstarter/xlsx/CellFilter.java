package org.jspring.dataexportstarter.xlsx;

import org.apache.poi.ss.usermodel.Cell;

import java.util.function.Predicate;

public enum CellFilter {
    NO_FILTER(cell -> true),
    FIRST_COLUMN(cell -> cell.getColumnIndex() == 0),
    FIRST_ROW(cell -> cell.getRowIndex() == 0);

    private final Predicate<Cell> filter;

    CellFilter(Predicate<Cell> filter) {
        this.filter = filter;
    }

    public Predicate<Cell> predicate() {
        return filter;
    }
}
