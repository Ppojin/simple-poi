package com.ppoijn.simplepoi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Util {
    public List<Object> rowToList(Row row){
        List<Object> rowList = new ArrayList<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case NUMERIC:
                    rowList.add(cell.getNumericCellValue());
                    break;
                case STRING:
                    rowList.add(cell.getStringCellValue());
                    break;
                case BOOLEAN:
                    rowList.add(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    rowList.add(cell.getCellFormula());
                    break;
                case BLANK:
                    rowList.add("");
                    break;
                case ERROR:
                    rowList.add(cell.getErrorCellValue());
                    break;
                default:
                    rowList.add(null);
                    break;
            }
        }
        return rowList;
    }

    public List<List<Object>> sheetToList(Sheet sheet) {
        List<List<Object>> excelData = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            excelData.add(rowToList(row));
        }
        return excelData;
    }

    public List<List<List<Object>>> getMergedSheetList(SimpleWorkbook simpleWorkbook) {
        List<List<List<Object>>> result = new ArrayList<>();

        Iterator<Sheet> sheetIterator = simpleWorkbook.getWorkbook().sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            result.add(sheetToList(sheet));
        }

        return result;
    }

    protected static CellStyle createCellStyle(SimpleWorkbook simpleWorkbook, IndexedColors color) {
        CellStyle newStyle = simpleWorkbook.getWorkbook().createCellStyle();
        newStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        newStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        newStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        newStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        newStyle.setBorderBottom(BorderStyle.THIN);
        newStyle.setBorderLeft(BorderStyle.THIN);
        newStyle.setBorderRight(BorderStyle.THIN);
        newStyle.setBorderTop(BorderStyle.THIN);
        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        newStyle.setFillForegroundColor(color.getIndex());
        newStyle.setAlignment(HorizontalAlignment.CENTER);
        return newStyle;
    }
}
