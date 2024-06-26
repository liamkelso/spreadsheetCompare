import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class SpreadsheetReader {

    private static final long MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB

    /**
     * Reads the spreadsheet and stores relevant data in a map.
     *
     * @param filePath    The path to the spreadsheet file.
     * @param idColumn    The name of the column that contains the employee IDs.
     * @param columnNames The names of the columns to compare.
     * @return A map where each key is an employee ID and the value is another map of column names to cell values.
     * @throws IOException If an I/O error occurs.
     */
    public static Map<String, Map<String, String>> readSpreadsheet(String filePath, String idColumn, String[] columnNames) throws IOException {
        File file = new File(filePath);

        // Check if the file is too large
        if (file.length() > MAX_FILE_SIZE) {
            System.err.println("File size exceeds the maximum limit of " + (MAX_FILE_SIZE / (1024 * 1024)) + " MB.");
            System.exit(1);
        }

        Map<String, Map<String, String>> dataMap = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Get the column indexes for the specified column names
            Map<String, Integer> columnIndexes = new HashMap<>();
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                String cellValue = cell.getStringCellValue().trim();
                if (cellValue.equalsIgnoreCase(idColumn)) {
                    columnIndexes.put("ID", cell.getColumnIndex());
                }
                for (String columnName : columnNames) {
                    if (columnName.equalsIgnoreCase(cellValue)) {
                        columnIndexes.put(columnName, cell.getColumnIndex());
                    }
                }
            }

            // Check if the ID column was found
            if (!columnIndexes.containsKey("ID")) {
                throw new IOException("ID column not found in the spreadsheet.");
            }

            // Check if all column names were found
            for (String columnName : columnNames) {
                if (!columnIndexes.containsKey(columnName)) {
                    throw new IOException("Column not found in the spreadsheet.");
                }
            }

            // Read the data rows
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                String id = getCellValue(row.getCell(columnIndexes.get("ID"))).trim();
                if (id.isEmpty()) {
                    System.out.println("Warning: Empty ID found in the spreadsheet, skipping row.");
                    continue;
                }

                Map<String, String> values = new HashMap<>();
                for (String columnName : columnNames) {
                    int colIndex = columnIndexes.get(columnName);
                    values.put(columnName, getCellValue(row.getCell(colIndex)).trim());
                }

                // Store the data in the map with ID as the key
                dataMap.put(id, values);
            }
        } catch (IOException e) {
            System.err.println("Error reading spreadsheet: " + filePath);
            System.exit(1);
        }

        return dataMap;
    }

    /**
     * Helper method to get the cell value as a string, handling different cell types.
     *
     * @param cell The cell to retrieve the value from.
     * @return The cell value as a string.
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}

