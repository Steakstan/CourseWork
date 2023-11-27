package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellType;

public class PostProcessing {

    private static int checkAndFillEmptyCellsInColumn(Sheet sheet, int columnIndex, int rowCount) {
        int lastRow = sheet.getLastRowNum();

        for (int rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
            Row currentRow = sheet.getRow(rowIndex);

            if (currentRow != null) {
                Cell currentCell = currentRow.getCell(columnIndex);

                if (currentCell == null || currentCell.getCellType() == CellType.BLANK) {
                    Cell cellAbove = sheet.getRow(rowIndex - 1).getCell(columnIndex);

                    if (cellAbove != null && cellAbove.getCellType() == CellType.STRING) {
                        if (currentCell == null) {
                            currentCell = currentRow.createCell(columnIndex);
                        }
                        currentCell.setCellValue(cellAbove.getStringCellValue());
                        rowCount++; // Increment rowCount after filling the empty cell
                    }
                }
            }
        }

        return rowCount;
    }

    public static int checkAndFillEmptyCellsInFirstColumn(Sheet sheet, int rowCount) {
        return checkAndFillEmptyCellsInColumn(sheet, 0, rowCount);
    }

    public static int checkAndFillEmptyCellsInSecondColumn(Sheet sheet, int rowCount) {
        return checkAndFillEmptyCellsInColumn(sheet, 1, rowCount);
    }

    public static int checkAndFillEmptyCellsInFourthColumn(Sheet sheet, int rowCount) {
        int lastRow = sheet.getLastRowNum();

        for (int rowIndex = lastRow; rowIndex >= 2; rowIndex--) {
            Row currentRow = sheet.getRow(rowIndex);

            if (currentRow != null) {
                Cell currentCell = currentRow.getCell(3);

                if (currentCell == null || currentCell.getCellType() == CellType.BLANK) {
                    Cell cellAbove = sheet.getRow(rowIndex + 1).getCell(3);

                    if (cellAbove != null && cellAbove.getCellType() == CellType.STRING) {
                        if (currentCell == null) {
                            currentCell = currentRow.createCell(3);
                        }
                        currentCell.setCellValue(cellAbove.getStringCellValue());
                        rowCount++; // Increment rowCount after filling the empty cell
                    }
                }
            }
        }

        return rowCount;
    }
}