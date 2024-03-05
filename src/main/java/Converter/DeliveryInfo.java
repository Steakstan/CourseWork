package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DeliveryInfo {

    public static void processDeliveryInfo(Sheet sheet, int rowCount, String line) {
        String extractedDeliveryDate = extractDeliveryDate(line);

        if (shouldProcessDeliveryInfo(extractedDeliveryDate, line)) {
            addDeliveryInfoToSheet(sheet, rowCount, extractedDeliveryDate);
        }
    }

    private static String extractDeliveryDate(String line) {
        Pattern deliveryPattern = Pattern.compile("KW\\s\\d{2}\\.\\d{4}|Auslauf|Neuanlauf|ersatlos|ausgelaufen|Bezugsberechtigung|nicht bezogen|Berechtigung|nicht bekannt|unbekannt");
        Matcher deliveryMatcher = deliveryPattern.matcher(line);

        String lastMatch = null;
        while (deliveryMatcher.find()) {
            lastMatch = deliveryMatcher.group();
        }
        return lastMatch;
    }

    private static boolean shouldProcessDeliveryInfo(String deliveryDate, String line) {
        return deliveryDate != null && !line.matches("(144|145)\\d{0,7}");
    }

    private static void addDeliveryInfoToSheet(Sheet sheet, int rowCount, String deliveryDate) {
        Row deliveryRow = sheet.getRow(rowCount - 1);
        if (deliveryRow == null) {
            deliveryRow = sheet.createRow(rowCount);
        }

        Cell deliveryCell = deliveryRow.createCell(3);
        deliveryCell.setCellValue(deliveryDate);
    }
}