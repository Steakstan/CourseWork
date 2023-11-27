package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.regex.Matcher;

public class BranchNumberProcessor {

    public static int processBranchNumberToExcel(String[] words, Sheet sheet, int rowCount) {
        String foundBranchNumber;

        for (int i = 0; i < words.length; i++) {
            foundBranchNumber = extractBranchNumber(words[i]);

            if (foundBranchNumber != null) {
                if (!checkOrderAfterBranch(words, i)) {
                    Row row = createRowIfNull(sheet, rowCount);
                    setCellValue(row, 1, foundBranchNumber);
                }
                break;
            }
        }

        return rowCount;
    }

    private static String extractBranchNumber(String word) {
        for (BranchPattern pattern : BranchPattern.values()) {
            Matcher matcher = pattern.getPattern().matcher(word);
            if (matcher.find()) {
                String foundBranchNumber = matcher.group();
                if (foundBranchNumber.length() == 5 || foundBranchNumber.length() == 6) {
                    return foundBranchNumber;
                }
            }
        }
        return null;
    }

    private static boolean checkOrderAfterBranch(String[] words, int startIndex) {
        boolean foundOrder = false;
        for (int i = startIndex + 1; i < words.length; i++) {
            if (words[i].matches("(144|145)\\d{0,7}")) {
                foundOrder = true;
                break;
            }
            if (words[i].equals(words[startIndex])) {
                break; // Break if the branch number repeats before finding the order number
            }
        }
        return foundOrder;
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
