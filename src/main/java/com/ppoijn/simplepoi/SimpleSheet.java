package com.ppoijn.simplepoi;

import lombok.Getter;
import lombok.Setter;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

public class SimpleSheet implements SimpleSheetInterface {
    @Getter
    private final String sheetName;
    private final Sheet sheet;

    private final CellStyle cellStyle;
    @Setter
    private boolean isAutoColumnSize = false;
    @Setter
    private int columnWidth = -1;

    private int nextRowNum = 0;
    private int nextColNum = 0;
    private int maxColLength = 0;
    private Row currentRow;

    public SimpleSheet(Sheet sheet) {
        this.sheet = sheet;
        this.sheetName = sheet.getSheetName();
        this.currentRow = sheet.getRow(this.nextRowNum);
        this.nextRowNum = sheet.getLastRowNum() + 1;
        if (this.currentRow != null) {
            this.nextColNum = this.currentRow.getLastCellNum() + 2;
        }
        this.cellStyle = null;
    }

    public SimpleSheet(Sheet sheet, String sheetName, List<Object> headerNames, CellStyle cellStyle) {
        this.sheetName = sheetName;
        this.sheet = sheet;
        this.cellStyle = cellStyle;
        this.nextRowNum = 0;
        this.currentRow = sheet.createRow(nextRowNum);
        if (CollectionUtils.isNotEmpty(headerNames)){
            this.insertRow(headerNames, cellStyle);
        }
    }

    @Override
    public void insertRowList(List<List<Object>> data) {
        data.forEach(this::insertRow);
    }

    @Override
    public void insertRow() {
        this.insertRow(new ArrayList<>());
    }

    @Override
    public void insertRow(List<Object> values) {
        this.insertRow(values, cellStyle);
    }

    @Override
    public void insertRow(List<Object> values, CellStyle style) {
        currentRow = sheet.createRow(nextRowNum);
        values.stream()
                .map(v -> ObjectUtils.defaultIfNull(v, ""))
                .forEach((v) -> insertCell(v, style));
        nextColNum = 0;
        nextRowNum++;
    }

    @Override
    public void insertCellList(List<Object> value) {
        value.forEach(this::insertCell);
    }

    @Override
    public void insertCell(Object value) {
        insertCell(value, cellStyle);
    }

    @Override
    public void insertCell(Object value, CellStyle style) {
        if (maxColLength < nextColNum) {
            maxColLength = nextColNum;
        }
        Cell cell = currentRow.createCell(nextColNum);
        cell.setCellValue(String.valueOf(value));
        if (!ObjectUtils.isEmpty(style)) {
            cell.setCellStyle(style);
        }
        nextColNum++;
    }

    public void done() {
        if (isAutoColumnSize) {
            IntStream.range(0, maxColLength)
                    .forEach(sheet::autoSizeColumn);
        } else if (columnWidth > -1) {
            IntStream.range(0, maxColLength)
                    .forEach((i) -> sheet.setColumnWidth(i, columnWidth));
        }
    }
}
