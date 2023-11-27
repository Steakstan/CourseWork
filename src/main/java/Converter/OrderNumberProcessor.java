package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class OrderNumberProcessor {

    public static int processOrderNumberToExcel(String[] words, Sheet sheet, int rowCount) {
        String foundOrderNumber = null;

        for (String word : words) {
            if (foundOrderNumber == null && isOrderNumber(word)) {
                foundOrderNumber = word;
                break; // Exit the loop if order number is found
            }
        }

        if (foundOrderNumber != null) {
            Row row = createRowIfNull(sheet, rowCount);
            setCellValue(row, 0, foundOrderNumber);
        }

        return rowCount;
    }

    private static boolean isOrderNumber(String word) {
        return word.matches("(144|145)\\d{0,7}");
    }

    private static Row createRowIfNull(Sheet sheet, int rowCount) {
        Row row = sheet.getRow(rowCount);
        return (row == null) ? sheet.createRow(rowCount) : row;
    }

    private static void setCellValue(Row row, int cellIndex, String value) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(value);
    }

}
