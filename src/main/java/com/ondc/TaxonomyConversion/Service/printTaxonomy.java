package com.ondc.TaxonomyConversion.Service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class printTaxonomy {

    private static Map<String, String> getTaxonomyMap() {
        Map<String, String> InputOutput = new HashMap<>();
        InputOutput.put("Autoparts & Components Taxonomy v2.0", "output-file-autoparts");
        InputOutput.put("Bulding & Construction Supplies v2.0", "output-file-construction");
        InputOutput.put("Fashion Taxonomy v2.0", "output-file-fashion");
        InputOutput.put("Chemical Taxonomy v2.0", "output-file-chemical");
        InputOutput.put("Hardware & Industrial Equipment v2.0", "output-file-hardware");
        InputOutput.put("Electronics & Electrical Appliances v2.0", "output-file-electronics");
        return InputOutput;
    }

    public static void main(String[] args) {
        Map<String, String> InputOutput = getTaxonomyMap();

        for (Map.Entry<String, String> s : InputOutput.entrySet()) {
            String excelFilePath = "C:\\Users\\Sonali Shakya\\Documents\\GitHub\\buyer-app\\TaxonomyConversion\\src\\main\\java\\com\\ondc\\TaxonomyConversion\\Taxonomy\\" + s.getKey() + ".xlsx";
            String outputFilePath = "C:\\Users\\Sonali Shakya\\Documents\\GitHub\\buyer-app\\TaxonomyConversion\\src\\main\\java\\com\\ondc\\TaxonomyConversion\\Output\\" + s.getValue() + ".json";

            System.out.println("Processing file - " + s.getKey());

            try {
                FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
                Workbook workbook = new XSSFWorkbook(excelFile);
                Sheet sheet = workbook.getSheetAt(0);

                // Map to hold JSON output
                Map<String, String> jsonData = new LinkedHashMap<>();

                // Get the header row (first row with column titles)
                Row headerRow = sheet.getRow(2);

//            System.out.println(headerRow.getCell(1).getStringCellValue());
                // Find the column index for "Code"
                int codeColumnIndex = findColumnIndex(headerRow, "Code");

                // Iterate over the rows starting from row 1 (after headers)
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);

                    if (row != null) {
                        // Get the value from the "Code" column
                        Cell codeCell = row.getCell(codeColumnIndex);
                        if (codeCell != null && codeCell.getCellType() == CellType.STRING) {
                            String code = codeCell.getStringCellValue();

                            // Creating JSON-like structure
                            StringBuilder jsonBuilder = new StringBuilder();
                            jsonBuilder.append("\"").append(code).append("\": {\n");

                            // Iterate over the remaining columns from the header row to get keys
                            for (int j = codeColumnIndex + 1; j < headerRow.getLastCellNum(); j++) {
                                Cell headerCell = headerRow.getCell(j);
                                Cell valueCell = row.getCell(j);

//                                System.out.println(valueCell);
//                                System.out.println("--------------");

                                if (headerCell != null && valueCell != null && valueCell.getCellType() == CellType.STRING) {
                                    String key = headerCell.getStringCellValue();
                                    String value = valueCell.getStringCellValue();

                                    // Clean the key according to the rules
                                    key = cleanKey(key);

                                    // Only process rows with "MC" or "MI" (mandatory true)
                                    String mandatory = getMandatoryValue(value);
                                    if ("true".equals(mandatory)) {
                                        // If key is "colour", handle it with the "COLOUR" value
                                        String keyValue = "colour".equalsIgnoreCase(key) ? "COLOUR" : "[]";

                                        jsonBuilder.append("  ").append(key).append(": { mandatory: ")
                                                .append(mandatory).append(", value: ").append(keyValue).append(" },\n");
                                    }
                                }
                            }

                            // End of the object
                            jsonBuilder.append("},\n");
                            jsonData.put(code, jsonBuilder.toString());
                        }
                    }
                }

                workbook.close();

                // Write the data to a text file
                BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath));

                for (String code : jsonData.keySet()) {
                    writer.write(jsonData.get(code));
                }

                writer.close();
                System.out.println("JSON data written to the output file.");

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    // Method to find column index based on header name
    private static int findColumnIndex(Row headerRow, String columnName) {
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
//            System.out.println(cell + "-" + columnName);
            if (cell != null && columnName.equalsIgnoreCase(cell.getStringCellValue().trim())) {
                return i;
            }
        }
        throw new IllegalArgumentException("Column with name '" + columnName + "' not found");
    }

    // Method to clean the key (replace spaces with underscores, remove brackets, handle slashes)
    private static String cleanKey(String key) {

        // Remove content in parentheses
        key = key.replaceAll("\\(.*?\\)", "");

        key = key.trim();

        // Replace spaces with underscores
        key = key.replaceAll("\\s+", "_");

        key = key.replaceAll("&", "And");

        // Remove everything after and including a slash
        if (key.contains("/")) {
            key = key.substring(0, key.indexOf("/"));
        }

        return key.trim(); // trim any extra spaces
    }

    // Method to return mandatory value based on cell content
    private static String getMandatoryValue(String cellValue) {
        if ("MC".equalsIgnoreCase(cellValue) || "MI".equalsIgnoreCase(cellValue)) {
            return "true";
        } else if ("O".equalsIgnoreCase(cellValue)) {
            return "false";
        }
        return "false";
    }
}
