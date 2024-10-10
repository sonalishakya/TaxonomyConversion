package com.ondc.TaxonomyConversion.Service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class printElectronics {

    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\Sonali Shakya\\Documents\\GitHub\\buyer-app\\TaxonomyConversion\\src\\main\\java\\com\\ondc\\TaxonomyConversion\\Taxonomy\\Electronics & Electrical Appliances v2.0.xlsx";
        String outputFilePath = "C:\\Users\\Sonali Shakya\\Documents\\GitHub\\buyer-app\\TaxonomyConversion\\src\\main\\java\\com\\ondc\\TaxonomyConversion\\Output\\output-file-electronics.txt";

        try {
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);

            // Map to hold JSON output
            Map<String, String> jsonData = new HashMap<>();

            // Get the header row (first row)
            Row headerRow = sheet.getRow(0);

            // Iterate over the rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                if (row != null) {
                    String code = row.getCell(3).getStringCellValue(); // assuming column index 3 is 'Code'

                    // Creating JSON-like structure
                    StringBuilder jsonBuilder = new StringBuilder();
                    jsonBuilder.append("\"").append(code).append("\": {\n");

                    jsonBuilder.append("  brand: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(4))).append(", value: [] },\n");

                    jsonBuilder.append("  model: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(5))).append(", value: [] },\n");

                    jsonBuilder.append("  model_year: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(6))).append(", value: \"/^[0-9]{1,4}$/\" },\n");

                    jsonBuilder.append("  colour: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(35))).append(", value: COLOUR },\n");

                    jsonBuilder.append("  colour_name: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(36))).append(", value: [] },\n");

                    jsonBuilder.append("  ram: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(10))).append(", value: \"/^[0-9]{1,3}$/\" },\n");

                    jsonBuilder.append("  ram_unit: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(11))).append(", value: [] },\n");

                    jsonBuilder.append("  rom: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(12))).append(", value: \"/^[0-9]{1,3}$/\" },\n");

                    jsonBuilder.append("  rom_unit: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(13))).append(", value: [] },\n");

                    jsonBuilder.append("  storage: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(14))).append(", value: \"/^[0-9]{1,4}$/\" },\n");

                    jsonBuilder.append("  storage_unit: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(15))).append(", value: [] },\n");

                    jsonBuilder.append("  primary_camera: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(18))).append(", value: \"/^[0-9]{1,3}$/\" },\n");

                    jsonBuilder.append("  secondary_camera: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(19))).append(", value: \"/^[0-9]{1,3}$/\" },\n");

                    jsonBuilder.append("  battery_capacity: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(20))).append(", value: \"/^[0-9]{1,5}$/\" },\n");

                    jsonBuilder.append("  os_type: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(22))).append(", value: [] },\n");

                    jsonBuilder.append("  os_version: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(23))).append(", value: [] },\n");

                    jsonBuilder.append("  weight: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(38))).append(", value: \"/^[0-9]+(.[0-9]{1,2})?$/\" },\n");

                    jsonBuilder.append("  length: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(40))).append(", value: \"/^[0-9]+(.[0-9]{1,2})?$/\" },\n");

                    jsonBuilder.append("  breadth: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(41))).append(", value: \"/^[0-9]+(.[0-9]{1,2})?$/\" },\n");

                    jsonBuilder.append("  height: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(42))).append(", value: \"/^[0-9]+(.[0-9]{1,2})?$/\" },\n");

                    jsonBuilder.append("  refurbished: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(37))).append(", value: [] },\n");

                    // Add more fields based on your payload structure...

                    // End of the object
                    jsonBuilder.append("},\n");

                    jsonData.put(code, jsonBuilder.toString());
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

    // Method to return mandatory value based on cell content
    private static String getMandatoryValue(Cell cell) {
        if (cell != null) {
            String cellValue = cell.getStringCellValue();
            if ("MC".equalsIgnoreCase(cellValue) || "MI".equalsIgnoreCase(cellValue)) {
                return "true";
            } else if ("O".equalsIgnoreCase(cellValue)) {
                return "false";
            }
        }
        return "false";
    }
}
