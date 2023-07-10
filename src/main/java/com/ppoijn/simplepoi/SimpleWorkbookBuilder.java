package com.ppoijn.simplepoi;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

public class SimpleWorkbookBuilder {
    public interface Builder {}

    public static class SheetNameBuilder {
        public Builder sheetName(String sheetName) {
            return sheetName(sheetName, false);
        }

        public Builder sheetName(String sheetName, boolean useHeaderRow) {
            if (useHeaderRow){
                return new HeaderNameBuilder(sheetName);
            } else {
                return new ElseBuilder(sheetName);
            }
        }

        public SimpleWorkbook excelFileInputStream(InputStream excelFileInputStream) throws SimplePoiException {
            try {
                return new SimpleWorkbook(excelFileInputStream);
            } catch (IOException e) {
                throw new SimplePoiException(e, Code.ETC, "");
            }
        }
    }

    public static class HeaderNameBuilder implements Builder {
        private final String sheetName;

        public HeaderNameBuilder(String sheetName) {
            this.sheetName = sheetName;
        }

        public ElseBuilder headerNames(String[] headerNames) {
            return headerNames(List.of(headerNames));
        }

        public ElseBuilder headerNames(List<String> headerNames) {
            return new ElseBuilder(
                    sheetName,
                    headerNames.stream()
                            .map(String::valueOf)
                            .collect(Collectors.toList())
            );
        }
    }

    public static class ElseBuilder implements Builder {
        private final String sheetName;
        private List<Object> headerNames;

        private int rowAccessWindowSize = 1000;
        private boolean isStyle = false;
        private String errorMessage = "something went wrong";
        private List<List<Object>> data;
        private Integer columnWidth;
        private boolean isAutoColumnSize = false;

        public ElseBuilder(String sheetName) {
            this.sheetName = sheetName;
        }

        public ElseBuilder(String sheetName, List<Object> headerNames) {
            this(sheetName);
            this.headerNames = headerNames;
        }

        public ElseBuilder rowAccessWindowSize(int rowAccessWindowsSize) {
            this.rowAccessWindowSize = rowAccessWindowsSize;
            return this;
        }

        public ElseBuilder style(boolean isStyle) {
            this.isStyle = isStyle;
            return this;
        }

        public ElseBuilder errorMessage(String errorMessage) {
            this.errorMessage = errorMessage;
            return this;
        }

        public ElseBuilder sheetData(List<List<Object>> data) {
            this.data = data;
            return this;
        }

        public ElseBuilder columnWidth(int columnWidth) {
            this.columnWidth = columnWidth;
            return this;
        }

        public ElseBuilder autoColumnSize(boolean isAutoColumnSize) {
            this.isAutoColumnSize = isAutoColumnSize;
            return this;
        }

        public SimpleWorkbook build() {
            SimpleWorkbook result = new SimpleWorkbook(sheetName, rowAccessWindowSize, isStyle, errorMessage, headerNames);
            SimpleSheet sheet = result.getCurrentSheet();
            if (Objects.nonNull(columnWidth)) {
                sheet.setColumnWidth(this.columnWidth);
            }
            sheet.setAutoColumnSize(this.isAutoColumnSize);
            if (Objects.nonNull(data)) {
                sheet.insertRowList(data);
            }
            return result;
        }
    }
}
