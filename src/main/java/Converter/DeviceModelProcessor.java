package Converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.regex.Matcher;

public class DeviceModelProcessor {

    public static int processModelsToExcel(String line, Sheet sheet, int rowCount) {
        for (DeviceModel device : DeviceModel.values()) {
            Matcher deviceMatcher = device.getPattern().matcher(line);

            if (deviceMatcher.find()) {
                Row deviceRow = sheet.getRow(rowCount);
                if (deviceRow == null) {
                    deviceRow = sheet.createRow(rowCount);
                }

                Cell deviceCell = deviceRow.createCell(2);
                deviceCell.setCellValue(device.name());
                rowCount++;
            }
        }
        return rowCount; // Return the updated rowCount
    }
}
