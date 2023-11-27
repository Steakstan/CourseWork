package Converter;
import org.apache.poi.ss.usermodel.*;

public class ExcelStyler {
    private static final int DEFAULT_COLUMN_WIDTH = 4000;

    private static CellStyle boldCenterStyle;
    private static CellStyle centerStyle;

    public static void setupHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"AB-Nummer", "Vertragsnummern", "Modellbezeichnung", "Liefertermin"};

        if (boldCenterStyle == null) {
            boldCenterStyle = createBoldCenterStyle(sheet);
        }
        if (centerStyle == null) {
            centerStyle = createCenterStyle(sheet);
        }

        for (int i = 0; i < headers.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headers[i]);
            headerCell.setCellStyle(boldCenterStyle);
            sheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH);
        }

        applyColumnStyle(sheet, centerStyle, headers.length);
    }

    private static CellStyle createBoldCenterStyle(Sheet sheet) {
        CellStyle boldCenterStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        boldCenterStyle.setFont(font);
        boldCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        return boldCenterStyle;
    }

    private static CellStyle createCenterStyle(Sheet sheet) {
        CellStyle centerStyle = sheet.getWorkbook().createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        return centerStyle;
    }

    private static void applyColumnStyle(Sheet sheet, CellStyle style, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.setDefaultColumnStyle(i, style);
        }
    }

    public static void removeEmptyRows(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        int i = lastRowNum;
        while (i >= 0) {
            Row row = sheet.getRow(i);
            if (isEmptyRow(row)) {
                sheet.shiftRows(i + 1, lastRowNum + 1, -1);
                lastRowNum--;
            } else {
                i--;
            }
        }
    }

    private static boolean isEmptyRow(Row row) {
        if (row == null) {
            return true;
        }
        for (int j = 0; j < row.getLastCellNum(); j++) {
            Cell cell = row.getCell(j);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }
    public static void addBordersToDataArea(Sheet sheet) {
        setBorders(sheet, sheet.getFirstRowNum(), sheet.getLastRowNum(), 0, sheet.getRow(sheet.getFirstRowNum()).getLastCellNum() - 1);
    }

    private static void setBorders(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        CellStyle borderedCellStyle = createBorderedCellStyle(sheet);

        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            boolean isRowEmpty = isEmptyRow(row);

            for (int colIndex = firstCol; colIndex <= lastCol; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }

                if (!isRowEmpty || (isRowEmpty && colIndex < row.getLastCellNum() && row.getCell(colIndex).getCellType() != CellType.BLANK)) {
                    cell.setCellStyle(borderedCellStyle);
                }
            }
        }
    }

    private static CellStyle createBorderedCellStyle(Sheet sheet) {
        CellStyle borderedCellStyle = sheet.getWorkbook().createCellStyle();
        borderedCellStyle.setBorderBottom(BorderStyle.THIN);
        borderedCellStyle.setBorderTop(BorderStyle.THIN);
        borderedCellStyle.setBorderLeft(BorderStyle.THIN);
        borderedCellStyle.setBorderRight(BorderStyle.THIN);
        return borderedCellStyle;
    }

}
