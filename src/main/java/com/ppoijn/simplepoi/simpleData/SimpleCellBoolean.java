package com.ppoijn.simplepoi.simpleData;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;

public class SimpleCellBoolean implements SimpleCell {
    private final CellType cellType = CellType.BOOLEAN;
    private final Boolean value;

    public SimpleCellBoolean(Boolean value) {
        this.value = value;
    }

    @Override
    public Boolean getValue() {
        return value;
    }
}
