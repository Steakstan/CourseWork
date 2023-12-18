package Converter;

import org.apache.poi.ss.usermodel.*;

public class ColumnCellChecker {
    private static int checkAndFillEmptyCellsInColumnFromBottom(Sheet sheet, int columnIndex, int rowCount) {
        int lastRow = sheet.getLastRowNum();

        for (int rowIndex = lastRow; rowIndex >= 2; rowIndex--) {
            Row currentRow = sheet.getRow(rowIndex);

            if (currentRow != null) {
                Cell currentCell = currentRow.getCell(columnIndex);

                if (currentCell == null || currentCell.getCellType() == CellType.BLANK) {
                    if (rowIndex + 1 <= lastRow) {
                        Row rowBelow = sheet.getRow(rowIndex + 1);
                        if (rowBelow != null) {
                            Cell cellBelow = rowBelow.getCell(columnIndex);
                            if (cellBelow != null && cellBelow.getCellType() == CellType.STRING) {
                                if (currentCell == null) {
                                    currentCell = currentRow.createCell(columnIndex);
                                }
                                currentCell.setCellValue(cellBelow.getStringCellValue());
                                rowCount++; // Increment rowCount after filling the empty cell
                            }
                        }
                    }
                }
            }
        }

        return rowCount;
    }

    public static int checkAndFillEmptyCellsInFirstColumnFromBottom(Sheet sheet, int rowCount) {
        return checkAndFillEmptyCellsInColumnFromBottom(sheet, 0, rowCount);
    }

    public static int checkAndFillEmptyCellsInSecondColumnFromBottom(Sheet sheet, int rowCount) {
        return checkAndFillEmptyCellsInColumnFromBottom(sheet, 1, rowCount);
    }
}
