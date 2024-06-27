package com.example.demo.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@RestController
public class ExcelController {

    @GetMapping("/read-excel")
    public String readExcel() {
        String excelFilePath = "D:/hash1.xlsx"; // Update this path
        StringBuilder data = new StringBuilder();

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int numColumns = headerRow.getPhysicalNumberOfCells();
            Map<Integer, String> headers = new HashMap<>();

            // Read headers
            for (int i = 0; i < numColumns; i++) {
                headers.put(i, headerRow.getCell(i).getStringCellValue());
            }

            // Read data rows and store in HashMaps
            Map<Integer, Map<String, Object>> excelData = new HashMap<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, Object> rowData = new HashMap<>();

                for (int colIndex = 0; colIndex < numColumns; colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    Object value;
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                value = cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                value = cell.getNumericCellValue();
                                break;
                            default:
                                value = "UNKNOWN";
                                break;
                        }
                    } else {
                        value = "UNKNOWN";
                    }
                    rowData.put(headers.get(colIndex), value);
                }
                excelData.put(rowIndex, rowData);
            }

            // Print the data
            excelData.forEach((rowNum, rowData) -> {
                data.append("Row ").append(rowNum).append(": ").append(rowData).append("\n");
            });

        } catch (IOException e) {
            e.printStackTrace();
            return "Error reading the Excel file.";
        }

        return data.toString();
    }
}
