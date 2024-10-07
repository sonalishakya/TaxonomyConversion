package com.ondc.TaxonomyConversion.Taxonomy;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.HashMap;
import java.util.Map;


public class printAppliance {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\Sonali Shakya\\Documents\\GitHub\\buyer-app\\TaxonomyConversion\\Electronics & Electrical Appliances v2.0.xlsx";
        String outputFilePath = "C:\\Users\\Sonali Shakya\\Documents\\GitHub\\buyer-app\\TaxonomyConversion\\output-file-appliance.txt";

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

                    jsonBuilder.append("  colour: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(35))).append(", value: COLOUR },\n");

                    jsonBuilder.append("  colour_name: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(36))).append(", value: [] },\n");

                    jsonBuilder.append("  type: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(29))).append(", value: [] },\n");

                    jsonBuilder.append("  special_feature: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(30))).append(", value: [] },\n");

                    jsonBuilder.append("  includes: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(28))).append(", value: [] },\n");

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

                    jsonBuilder.append("  energy_rating: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(43))).append(", value: [] },\n");

                    jsonBuilder.append("  battery: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(44))).append(", value: [] },\n");

                    jsonBuilder.append("  power_input: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(45))).append(", value: [] },\n");

                    jsonBuilder.append("  warranty: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(46))).append(", value: [] },\n");

                    jsonBuilder.append("  extended_warranty: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(47))).append(", value: [] },\n");

                    jsonBuilder.append("  installation_detail: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(48))).append(", value: [] },\n");

                    jsonBuilder.append("  wattage: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(49))).append(", value: [] },\n");

                    jsonBuilder.append("  voltage: { mandatory: ")
                            .append(getMandatoryValue(row.getCell(50))).append(", value: [] },\n");

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
