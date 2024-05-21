package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class BranchNumberProcessor {
    public static String processBranchNumberToExcel(String[] words, Sheet sheet, int rowCount) {
        for (int i = 0; i < words.length; i++) {
            String foundBranchNumber = extractBranchNumber(words[i]);
            if (foundBranchNumber != null) {
                if (!checkOrderAfterBranch(words, i)) {
                    Row row = createRowIfNull(sheet, rowCount);
                    setCellValue(row, 1, foundBranchNumber);
                    return foundBranchNumber;
                }
                break;
            }
        }
        return null;
    }

    private static String extractBranchNumber(String word) {
        for (String branchNumber : BranchNumbers.BRANCH_NUMBERS) {
            if (word.matches("\\b" + branchNumber + "\\w{3,4}\\b")) {
                if ((word.length() == 5 || word.length() == 6) &&
                        !word.equalsIgnoreCase("120151") &&
                        !word.equalsIgnoreCase("ESSEN") &&
                        !word.equalsIgnoreCase("WILLI") &&
                        !word.equalsIgnoreCase("AUTARK") &&
                        !word.equalsIgnoreCase("ANZING") &&
                        !word.equalsIgnoreCase("FROST") &&
                        !word.equalsIgnoreCase("BRAND") &&
                        !word.equalsIgnoreCase("PALLEN") &&
                        !word.equalsIgnoreCase("12529") &&
                        !word.equalsIgnoreCase("16727") &&
                        !word.equalsIgnoreCase("BOSCH") &&
                        !word.equalsIgnoreCase("CENTER") &&
                        !word.equalsIgnoreCase("CLEAN") &&
                        !word.equalsIgnoreCase("PASSAU") &&
                        !word.equalsIgnoreCase("14513") &&
                        !word.equalsIgnoreCase("HALLE") &&
                        !word.equalsIgnoreCase("12439") &&
                        !word.equalsIgnoreCase("17033") &&
                        !word.equalsIgnoreCase("BLACK") &&
                        !word.equalsIgnoreCase("BRONZE") &&
                        !word.equalsIgnoreCase("SILVER") &&
                        !word.equalsIgnoreCase("12687") &&
                        !word.equalsIgnoreCase("PAMPOW")) {
                    return word;
                }
            }
        }
        // Consider returning a specific value or handling this case differently
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
