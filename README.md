# Taxonomy Excel to JSON Converter

This utility is designed to convert a taxonomy Excel file to a JSON format, to be used in log verification utilities for validating mandatory attributes.

## How to Use

1. **Clone the code** to your local system.
2. Format your Excel file to have the column header "Code" for domain codes (e.g., RET10-1004).
3. Copy your Excel file to the path:
   ```
   TaxonomyConversion\src\main\java\com\ondc\TaxonomyConversion\Taxonomy
   ```
4. In the `printTaxonomy` class at:
   ```
   TaxonomyConversion\src\main\java\com\ondc\TaxonomyConversion\Service
   ```
   In the `getTaxonomyMap` function, add your file name as the map key and the desired output file name as the value.
5. Set `excelFilePath` and `outputFilePath` relative to your local system's root.

6. **Run the code** and find the output files at:
   ```
   TaxonomyConversion\src\main\java\com\ondc\TaxonomyConversion\Output
   ```
