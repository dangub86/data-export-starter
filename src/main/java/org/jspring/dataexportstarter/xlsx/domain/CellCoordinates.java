package org.jspring.dataexportstarter.xlsx.domain;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.jspring.dataexportstarter.xlsx.CellFilter;

import java.util.function.Predicate;

public record CellCoordinates<T>(
        int rowNumber,
        int columnNumber,
        T cellValue,
        Predicate<Cell> filter
) {

    public boolean byCoordinates() {
        return rowNumber != -1 && columnNumber != -1;
    }

    public boolean byValue() {
        return cellValue != null;
    }

    /*
    STATIC BUILDER
     */
    public static class SearchBuilder<T> {
        private int rowNumber = -1;
        private int columnNumber = -1;
        private T cellValue;
        private Predicate<Cell> filter = CellFilter.NO_FILTER.predicate();

        public static <T> SearchBuilder<T> init() {
            return new SearchBuilder<>();
        }

        public SearchBuilder<T> rowNumber(int rowNumber) {
            this.rowNumber = rowNumber;
            return this;
        }

        public SearchBuilder<T> columnNumber(int columnNumber) {
            this.columnNumber = columnNumber;
            return this;
        }

        public SearchBuilder<T> cellValue(T cellValue) {
            this.cellValue = cellValue;
            return this;
        }

        public SearchBuilder<T> filter(Predicate<Cell> filter) {
            this.filter = filter;
            return this;
        }

        public SearchBuilder<T> firstColumn() {
            this.filter = CellFilter.FIRST_COLUMN.predicate();
            return this;
        }

        public SearchBuilder<T> firstRow() {
            this.filter = CellFilter.FIRST_ROW.predicate();
            return this;
        }

        public SearchBuilder<T> address(String address) {
            CellReference reference = new CellReference(address);
            this.rowNumber = reference.getRow();
            this.columnNumber = reference.getCol();
            return this;
        }

        public CellCoordinates<T> build() {
            return new CellCoordinates<>(rowNumber, columnNumber, cellValue, filter);
        }

    }

}
