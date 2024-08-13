package Converter;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.IOException;
public class ProcessPDFFile {

    static int processPDFFile(File pdfFile, Sheet sheet, int rowCount) {
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String extractedText = pdfStripper.getText(document);

            String[] extractedLines = extractedText.split("\n");

            for (String extractedLine : extractedLines) {
                String[] words = extractedLine.split("\\s+");

                // Process order and branch numbers

                OrderNumberProcessor.processOrderNumberToExcel(words, sheet, rowCount);
                BranchNumberProcessor.processBranchNumberToExcel(words, sheet, rowCount);


                // Process device models
                rowCount = DeviceModelProcessor.processModelsToExcel(extractedLine, sheet, rowCount);

                // Process delivery info
                DeliveryInfo.processDeliveryInfo(sheet, rowCount, extractedLine);

            }

          /*PostProcessing.checkAndFillEmptyCellsInFirstColumn(sheet, rowCount);
            PostProcessing.checkAndFillEmptyCellsInSecondColumn(sheet, rowCount);
            PostProcessing.checkAndFillEmptyCellsInFourthColumn(sheet, rowCount);*/



        } catch (IOException e) {
            // Handle PDF processing or Excel writing exceptions
            e.printStackTrace();
        }

        return rowCount;
    }
}