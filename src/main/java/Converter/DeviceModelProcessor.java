package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DeviceModelProcessor {
    public static int processModelsToExcel(String line, Sheet sheet, int rowCount) {
        String[] deviceModels = DeviceModels.getDeviceModels(); // Получаем массив моделей из отдельного класса

        for (String model : deviceModels) {
            Matcher deviceMatcher = Pattern.compile("\\b" + model + "\\b").matcher(line);

            if (deviceMatcher.find()) {
                Row deviceRow = sheet.getRow(rowCount);
                if (deviceRow == null) {
                    deviceRow = sheet.createRow(rowCount);
                }

                Cell deviceCell = deviceRow.createCell(2);
                deviceCell.setCellValue(model);
                rowCount++;
            }
        }
        return rowCount; // Return the updated rowCount
    }
}
