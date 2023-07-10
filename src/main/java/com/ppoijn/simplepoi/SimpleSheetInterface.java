package com.ppoijn.simplepoi;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;

public interface SimpleSheetInterface {
    void insertRowList(List<List<Object>> data);

    void insertRow();
    void insertRow(List<Object> values);
    void insertRow(List<Object> values, CellStyle style);

    void insertCellList(List<Object> value);
    void insertCell(Object value);
    void insertCell(Object value, CellStyle style);
}
