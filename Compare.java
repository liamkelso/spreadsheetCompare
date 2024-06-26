import java.io.File;
import java.io.IOException;
import java.util.*;

/**
 * This program compares data between two Excel spreadsheets based on a specified employee ID column.
 * It allows the user to specify which columns to compare, even if the column names are different between the spreadsheets.
 * 
 * The program performs the following steps:
 * 1. Prompts the user for the paths to the two spreadsheets.
 * 2. Prompts the user for the column names where the employee IDs are located in both spreadsheets.
 * 3. Prompts the user to specify if the columns to compare have the same names in both spreadsheets.
 * 4. Prompts the user for the names of the columns to compare.
 * 5. Reads data from both spreadsheets and stores it in maps with the employee ID as the key.
 * 6. Compares the specified columns for each employee ID.
 * 7. Prints any discrepancies found during the comparison.
 * 8. Prints "All information matches" if no discrepancies are found.
 * 
 * The program handles various edge cases such as missing columns, empty IDs, and mismatched column names.
 * 
 * Date last updated: 2024-06-26
 * 
 * @author liamkelso
 * @version 3.0
 */
public class Compare {

    private static final long MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // Get file paths for the two spreadsheets to be compared
        String file1 = getFilePath(scanner, "Enter the path for the first spreadsheet (e.g., spreadsheets/spreadsheet1.xlsx): ");
        String file2 = getFilePath(scanner, "Enter the path for the second spreadsheet (e.g., spreadsheets/spreadsheet2.xlsx): ");

        System.out.println("Please choose what column you want to use as the key. \nThis should be a constant between both spreadsheets, "
        		+ "something such as an employee ID or the policy numbeer.");
        
        // Ask for the column names where the employee IDs are located
        System.out.print("Enter the name of the column containing the keys you wish to use in the first spreadsheet: ");
        String idColumn1 = scanner.nextLine().trim();
        System.out.print("Enter the name of the column containing the keys you wish to use in the second spreadsheet: ");
        String idColumn2 = scanner.nextLine().trim();

        // Ask if the names of the columns are the same on both spreadsheets
        System.out.print("Are the names of the columns the same on both spreadsheets? (yes/no): ");
        String columnInfo = scanner.nextLine().trim();

        // Only allows user to enter yes or no, repeats prompt until one is entered
        while (!columnInfo.equalsIgnoreCase("yes") && (!columnInfo.equalsIgnoreCase("no"))) {
            System.out.println("Please enter yes or no.");
            System.out.print("Are the names of the columns the same on both spreadsheets? (yes/no): ");
            columnInfo = scanner.nextLine().trim();
        }

        // Get the number of columns to compare
        int numColumns = getNumberOfColumns(scanner);

        // Variables to store column names
        String[] columnsToCompare1;
        String[] columnsToCompare2;

        if (columnInfo.equalsIgnoreCase("yes")) {
            // Get the names of the columns to compare for both spreadsheets (same names)
            columnsToCompare1 = new String[numColumns];
            columnsToCompare2 = new String[numColumns];
            for (int i = 0; i < numColumns; i++) {
                System.out.print("Enter the name of column " + (i + 1) + " to compare: ");
                String columnName = scanner.nextLine().trim();
                columnsToCompare1[i] = columnName;
                columnsToCompare2[i] = columnName;
            }
        } else {
            // Get the names of the columns to compare for both spreadsheets (different names)
            columnsToCompare1 = new String[numColumns];
            columnsToCompare2 = new String[numColumns];
            for (int i = 0; i < numColumns; i++) {
                System.out.print("Enter the name of column " + (i + 1) + " you want to compare in the first spreadsheet: ");
                columnsToCompare1[i] = scanner.nextLine().trim();
                System.out.print("Enter the name of column " + (i + 1) + " you want to compare in the second spreadsheet: ");
                columnsToCompare2[i] = scanner.nextLine().trim();
            }
        }

        System.out.println("Please ignore the StatusLogger Error.");

        try {
            // Check file sizes
            checkFileSize(file1);
            checkFileSize(file2);

            // Read data from both spreadsheets
            Map<String, Map<String, String>> data1 = SpreadsheetReader.readSpreadsheet(file1, idColumn1, columnsToCompare1);
            Map<String, Map<String, String>> data2 = SpreadsheetReader.readSpreadsheet(file2, idColumn2, columnsToCompare2);

            // Compare the data from both spreadsheets
            compareSpreadsheets(data1, data2, columnsToCompare1, columnsToCompare2);

        } catch (IOException e) {
            System.err.println("Error reading spreadsheets: Please check the file paths and formats.");
        } catch (Exception e) {
            System.err.println("An unexpected error occurred. Please try again.");
        } finally {
            scanner.close();
        }
    }

    /**
     * Checks the file size and exits if it exceeds the maximum allowed size.
     *
     * @param filePath The path to the file.
     */
    private static void checkFileSize(String filePath) {
        File file = new File(filePath);
        if (file.length() > MAX_FILE_SIZE) {
            System.err.println("File size exceeds the maximum limit of " + (MAX_FILE_SIZE / (1024 * 1024)) + " MB.");
            System.exit(1);
        }
    }

    /**
     * Prompts the user to enter the file path.
     * 
     * @param scanner The Scanner object to read user input.
     * @param prompt  The message to prompt the user.
     * @return The file path entered by the user.
     */
    private static String getFilePath(Scanner scanner, String prompt) {
        String filePath;
        while (true) {
            System.out.print(prompt);
            filePath = scanner.nextLine().trim();
            if (!filePath.isEmpty()) {
                break;
            } else {
                System.err.println("File path cannot be empty. Please try again.");
            }
        }
        return filePath;
    }

    /**
     * Prompts the user to enter the number of columns to compare.
     * 
     * @param scanner The Scanner object to read user input.
     * @return The number of columns to compare.
     */
    private static int getNumberOfColumns(Scanner scanner) {
        int numColumns;
        while (true) {
            System.out.print("Enter the number of columns to compare: ");
            try {
                numColumns = Integer.parseInt(scanner.nextLine().trim());
                if (numColumns > 0) {
                    break;
                } else {
                    System.err.println("Number of columns must be greater than 0. Please try again.");
                }
            } catch (NumberFormatException e) {
                System.err.println("Invalid number. Please enter a valid integer.");
            }
        }
        return numColumns;
    }

    /**
     * Method to compare the data from two spreadsheets.
     * 
     * @param data1             The data from the first spreadsheet.
     * @param data2             The data from the second spreadsheet.
     * @param columnsToCompare1 The names of the columns to compare in the first
     *                          spreadsheet.
     * @param columnsToCompare2 The names of the columns to compare in the second
     *                          spreadsheet.
     */
    private static void compareSpreadsheets(Map<String, Map<String, String>> data1,
                                            Map<String, Map<String, String>> data2, String[] columnsToCompare1, String[] columnsToCompare2) {
        boolean allMatch = true; // Variable to track if all data matches

        for (String id : data1.keySet()) {
            Map<String, String> values1 = data1.get(id);
            Map<String, String> values2 = data2.get(id);

            if (values2 == null) {
                System.out.println("ID " + id + " is missing in the second spreadsheet.");
                allMatch = false;
                continue;
            }

            boolean match = true;
            for (int i = 0; i < columnsToCompare1.length; i++) {
                String column1 = columnsToCompare1[i].trim();
                String column2 = columnsToCompare2[i].trim();
                if (!values1.get(column1).equals(values2.get(column2))) {
                    match = false;
                    allMatch = false;
                    break;
                }
            }

            if (!match) {
                System.out.println("Mismatch found for ID " + id + ":");
                for (int i = 0; i < columnsToCompare1.length; i++) {
                    String column1 = columnsToCompare1[i].trim();
                    String column2 = columnsToCompare2[i].trim();
                    System.out.println("  " + column1 + " vs " + column2 + ": " + values1.get(column1) + " vs " + values2.get(column2));
                }
            }
        }

        for (String id : data2.keySet()) {
            if (!data1.containsKey(id)) {
                System.out.println("ID " + id + " is missing in the first spreadsheet.");
                allMatch = false;
            }
        }

        if (allMatch) {
            System.out.println("All information matches.");
        }
    }
}
