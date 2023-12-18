package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class OrderInfoChecker {
    public static void checkOrderInformation(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell orderCell = row.getCell(0); // Получаем ячейку с номером заказа
                if (orderCell != null && orderCell.getCellType() == CellType.STRING) {
                    String orderNumber = orderCell.getStringCellValue();
                    if (isOrderNumberValid(orderNumber)) {
                        String branchNumber = getBranchNumberFromAdjacentCell(row);
                        if (branchNumber != null) {
                            int searchIndex = i + 1;
                            while (searchIndex <= lastRowNum) {
                                Row searchRow = sheet.getRow(searchIndex);
                                if (searchRow != null) {
                                    Cell searchOrderCell = searchRow.getCell(0);
                                    if (searchOrderCell != null && searchOrderCell.getCellType() == CellType.STRING) {
                                        String searchOrderNumber = searchOrderCell.getStringCellValue();
                                        if (orderNumber.equals(searchOrderNumber)) {
                                            Cell branchCell = searchRow.getCell(1); // Ячейка с номером филиала
                                            if (branchCell != null && branchCell.getCellType() == CellType.BLANK) {
                                                setBranchCellValue(searchRow, 1, branchNumber); // Копируем номер филиала
                                            }
                                        }
                                    }
                                }
                                searchIndex++;
                            }
                        }
                    }
                }
            }
        }
    }

    private static boolean isOrderNumberValid(String orderNumber) {
        return orderNumber.matches("(144|145)\\d{0,9}");
    }

    private static String getBranchNumberFromAdjacentCell(Row row) {
        Cell adjacentCell = row.getCell(1); // Ячейка с номером филиала
        if (adjacentCell != null && adjacentCell.getCellType() == CellType.STRING) {
            return adjacentCell.getStringCellValue();
        }
        return null;
    }

    private static void setBranchCellValue(Row row, int cellIndex, String value) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(value);
    }
}