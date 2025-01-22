package excelConvertor;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;

public class ReadExcel {

    public String getExcelData(String filePath, String sheetName, int[] cols) {
        StringBuilder excelData = new StringBuilder();
        try (FileInputStream fis = new FileInputStream(filePath)) { // Use try-with-resources to auto-close the stream
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sh = workbook.getSheet(sheetName);

            // If the sheet is not found, use the first sheet
            if (sh == null) {
                sh = workbook.getSheetAt(0);
            }

            int rows = sh.getLastRowNum();

            excelData.append("[\n");

            // Date formatter
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

            // Iterate through rows (skip header row)
            for (int i = 1; i <= rows; i++) {
                Row row = sh.getRow(i);
                if (row == null) {
                    continue; // Skip empty rows
                }

                excelData.append("{\n");

                // Iterate through specified columns
                for (int j = 0; j < cols.length; j++) {
                    int colIndex = cols[j]; // Get the actual column index
                    Cell cell = row.getCell(colIndex);

                    // Append the column header
                    excelData.append("\t\"");
                    excelData.append(sh.getRow(0).getCell(colIndex).toString()); // Header row
                    excelData.append("\" : ");

                    // Handle cell value based on its type
                    if (cell == null) {
                        excelData.append("null"); // Handle null cells
                    } else {
                        CellType ct = cell.getCellType();

                        switch (ct) {
                            case NUMERIC -> {
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    // Handle dates
                                    excelData.append("\"").append(dateFormat.format(cell.getDateCellValue())).append("\"");
                                } else {
                                    // Handle numeric values
                                    excelData.append(cell.getNumericCellValue());
                                }
                            }
                            case STRING -> // Handle string values
                                    excelData.append("\"").append(cell.getStringCellValue()).append("\"");
                            case BOOLEAN -> // Handle boolean values
                                    excelData.append(cell.getBooleanCellValue());
                            case FORMULA -> // Handle formula cells
                                    excelData.append("\"").append(cell.getCellFormula()).append("\"");
                            case BLANK -> // Handle blank cells
                                    excelData.append("null");
                            default -> // Handle other types (e.g., ERROR)
                                    excelData.append("\"UNKNOWN\"");
                        }

                    }

                    // Add a comma if it's not the last column
                    if (j < cols.length - 1) {
                        excelData.append(",\n");
                    } else {
                        excelData.append("\n");
                    }
                }

                // Add a comma if it's not the last row
                if (i < rows) {
                    excelData.append("},\n");
                } else {
                    excelData.append("}\n");
                }
            }

            excelData.append("]");
            System.out.println(rows + " rows converted successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }

        return excelData.toString();
    }
}