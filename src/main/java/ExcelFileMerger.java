import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.util.*;

public class ExcelFileMerger {
    public static void main(String[] args) {
        //Replace this by the various file paths.
        String[] fileNames = {"D:/JU ACADEMIC DOCUMENTS/Fourth Year/Sem 2/Project/SampleExcelFiles/Sheet1.xlsx", "D:/JU ACADEMIC DOCUMENTS/Fourth Year/Sem 2/Project/SampleExcelFiles/Sheet2.xlsx", "D:/JU ACADEMIC DOCUMENTS/Fourth Year/Sem 2/Project/SampleExcelFiles/Sheet3.xlsx"};
        Map<Integer, List<VoltageCurrent>> dataMap = new HashMap<>(); // Excel sheet no , (Potential,current) pair stored here.

        int excelSheetNumber = 0;
        for (String fileName : fileNames) {
            try (FileInputStream fis = new FileInputStream(fileName);
                 Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

                // Get column indices based on header names (adjust if needed)
                int potentialColumnIndex = sheet.getRow(0).getCell(0).getColumnIndex();
                int currentColumnIndex = sheet.getRow(0).getCell(1).getColumnIndex();
                dataMap.put(excelSheetNumber, new ArrayList<>());
                // Read data into a map
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    String potential = row.getCell(potentialColumnIndex).toString();
                    String current = row.getCell(currentColumnIndex).toString();
                    dataMap.get(excelSheetNumber).add(new VoltageCurrent(potential, current));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            excelSheetNumber++;
        }

        System.out.println("Data Stored in map from the Excel Sheets to be Merged " + dataMap);
        int maxNumberOfRows = dataMap.get(0).size();
        int numberOfSheetsToMerge = dataMap.size();
//       System.out.println(maxNumberOfRows +"  "+numberOfSheetsToMerge);


        // Resulting merged xlsx document will be Stored as per the file path provided.
        try (Workbook mergedWorkbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream("D:/JU ACADEMIC DOCUMENTS/Fourth Year/Sem 2/Project/SampleExcelFiles/merged_file.xlsx")) {
            Sheet mergedSheet = mergedWorkbook.createSheet("Merged Data");

            // Create header row
            Row headerRow = mergedSheet.createRow(0);
            headerRow.createCell(0).setCellValue("Potential");
            headerRow.createCell(1).setCellValue("Current1");
            headerRow.createCell(2).setCellValue("Current2");
            headerRow.createCell(3).setCellValue("Current3");

            // Populate data rows
            for (int currRow = 1; currRow <= maxNumberOfRows; currRow++) {
                Row dataRow = mergedSheet.createRow(currRow);
                for (int currColumn = 0; currColumn <= numberOfSheetsToMerge; currColumn++) {
                    if (currColumn == 0) {
                        dataRow.createCell(currColumn).setCellValue(dataMap.get(0).get(currRow - 1).voltage);
                    } else {
                        dataRow.createCell(currColumn).setCellValue(dataMap.get(currColumn - 1).get(currRow - 1).current);
                    }
                }
            }
            mergedWorkbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    // Class used to store the Potential Current Pairs.
    static class VoltageCurrent {
        String voltage;
        String current;

        public VoltageCurrent(String a, String b) {
            this.voltage = a;
            this.current = b;
        }

        @Override
        public String toString() {
            return "Voltage : " + voltage + " current : " + current;
        }
    }

}
