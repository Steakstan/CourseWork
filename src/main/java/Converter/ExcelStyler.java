package Converter;

import org.apache.poi.ss.usermodel.*;

public class ExcelStyler {
    // Method to style the first row and auto-fit columns based on content
    public static void styleFirstRowAndAutofitColumns(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"AB-Nummer", "Vertragsnummern", "Modellbezeichnung", "Liefertermin"};

        // Create style for bold and centered text
        CellStyle boldCenterStyle = createBoldCenterStyle(sheet);

        // Populate the first row and auto-fit columns based on content
        for (int i = 0; i < headers.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headers[i]);
            headerCell.setCellStyle(boldCenterStyle);
            sheet.autoSizeColumn(i); // Auto-fit column width based on content
        }
        headerRow.setHeightInPoints(20); // Set row height
    }

    // Method to create style for bold and centered text
    private static CellStyle createBoldCenterStyle(Sheet sheet) {
        CellStyle boldCenterStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        boldCenterStyle.setFont(font);
        boldCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        return boldCenterStyle;
    }

    // Method to add borders and align text in the first four columns
    public static void addBordersAndAlignTextInFirstFourColumns(Sheet sheet) {
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        int firstCol = 0;
        int lastCol = 3; // Four columns, starting from 0

        // Create style for bordered and centered text
        CellStyle borderedCenterCellStyle = createBorderedCenterCellStyle(sheet);

        // Iterate through rows and columns to apply styles
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            for (int colIndex = firstCol; colIndex <= lastCol; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }

                // Apply style if cell is not blank
                if (cell.getCellType() != CellType.BLANK) {
                    cell.setCellStyle(borderedCenterCellStyle);
                }
            }
        }
    }

    // Method to create style with borders and centered text
    private static CellStyle createBorderedCenterCellStyle(Sheet sheet) {
        CellStyle borderedCenterCellStyle = sheet.getWorkbook().createCellStyle();
        borderedCenterCellStyle.setBorderBottom(BorderStyle.THIN);
        borderedCenterCellStyle.setBorderTop(BorderStyle.THIN);
        borderedCenterCellStyle.setBorderLeft(BorderStyle.THIN);
        borderedCenterCellStyle.setBorderRight(BorderStyle.THIN);
        borderedCenterCellStyle.setAlignment(HorizontalAlignment.CENTER);
        return borderedCenterCellStyle;
    }
}