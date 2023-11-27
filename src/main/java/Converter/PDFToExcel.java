package Converter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static Converter.ProcessPDFFile.processPDFFile;

public class PDFToExcel {

    public static void main(String[] args) {
        // Path to the directory containing the PDF files
        String directoryPath = "/Users/user/Documents/AB/";

        // Path to the Excel file that will be created
        String excelFilePath = "/Users/user/Documents/AB-Nummern.xlsx";

        // Retrieve files from the specified directory
        File directory = new File(directoryPath);
        File[] files = directory.listFiles();

        // Create a new Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Order Numbers");
        int rowCount = 1; // Common rowCount for all files

        // Loop through the files in the directory
        if (files != null) {
            for (File file : files) {
                // Process only PDF files
                if (file.isFile() && (file.getName().endsWith(".pdf") || file.getName().endsWith(".PDF"))) {
                    rowCount = processPDFFile(file, sheet, rowCount);// Use and update the common rowCount

                }
            }
        }





        ExcelStyler.addBordersAndAlignTextInFirstFourColumns(sheet);
        // Set up the header row in the Excel sheet
        ExcelStyler.styleFirstRowAndAutofitColumns(sheet);
        // Write the workbook content to an Excel file
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
            System.out.println("Excel file created successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}