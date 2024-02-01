package org.jspring.dataexportstarter.xlsx.domain;

import org.apache.poi.ss.usermodel.CellType;

public record CellWrapper<T>(CellType cellType, T value) {
}
