package com.example.demo.excelreader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {

    public static void main(String[] args) {
       
//        if (args.length < 1) {
//            System.out.println("Usage: java ExcelReader <ExcelFilePath>");
//            return;
//        }


        String excelFilePath = "D:/hash1.xlsx";
        for (String arg :
        	args) {
            if (!arg.startsWith("--")) {
                excelFilePath = arg;
                break;
            }
        }


        
        File excelFile = new File(excelFilePath);
        if (!excelFile.exists() || !excelFile.isFile()) {
            System.err.println("The file does not exist or is not a valid file: " + excelFilePath);
            return;
        }

        Map<Integer, List<Object>> excelData = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(excelFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); 
          
            for (Row row : sheet) {
                List<Object> rowData = new ArrayList<>();

                for (Cell cell : row) {
                    rowData.add(getCellValue(cell));
                }

                
                excelData.put(row.getRowNum(), rowData);
            }

            excelData.forEach((rowNum, data) -> {
                System.out.println("Row " + rowNum + ": " + data);
            });

        } catch (IOException e) {
            System.err.println("Error reading the Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }

  
    private static Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                // Evaluate the formula and get the result
                FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                return getCellValue(evaluator.evaluateInCell(cell));
            case BLANK:
                return "";
            case ERROR:
                return "ERROR";
            default:
                return "UNKNOWN";
        }
    }
}