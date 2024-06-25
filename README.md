# Spreadsheet Comparison Tool

## Project Overview

This project is a Java-based tool for comparing data between two Excel spreadsheets. The comparison is based on a specified employee ID column, allowing the user to specify which columns to compare, even if the column names are different between the spreadsheets. 

The tool performs the following steps:
1. Prompts the user for the paths to the two spreadsheets.
2. Prompts the user for the column names where the employee IDs are located in both spreadsheets.
3. Prompts the user to specify if the columns to compare have the same names in both spreadsheets.
4. Prompts the user for the names of the columns to compare.
5. Reads data from both spreadsheets and stores it in maps with the employee ID as the key.
6. Compares the specified columns for each employee ID.
7. Prints any discrepancies found during the comparison.
8. Prints a statement "All information matches" if no discrepancies are found.

The program handles various edge cases such as missing columns, empty IDs, and mismatched column names.

## Author

- **Author**: liamkelso
- **Version**: 2.0
- **Date last updated**: 2024-06-25

## Prerequisites

- Java Development Kit (JDK) 8 or later
- Apache POI library for handling Excel files

## Setup Instructions

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/liamkelso/spreadsheet-comparison-tool.git
   cd spreadsheet-comparison-tool

2. **Install Apache POI Library**:

Download the Apache POI library from Apache POI Downloads.
Extract the downloaded files.
Add the following JAR files to your project's classpath:
poi-5.2.2.jar
poi-ooxml-5.2.2.jar
poi-ooxml-schemas-5.2.2.jar
xmlbeans-5.1.1.jar
commons-compress-1.21.jar
log4j-api-2.14.1.jar
log4j-core-2.14.1.jar

3. **Build and Run the Project:**

Open the project in your preferred IDE (e.g., Eclipse, IntelliJ).
Build the project to resolve dependencies.
Run the Compare class as a Java application.

## Usage Instructions

**Run the Program:**

Execute the Compare class to start the program.

**Follow the Prompts:**

Enter the paths for the two spreadsheets when prompted.
Enter the column names where the employee IDs are located in both spreadsheets.
Specify whether the columns to compare have the same names in both spreadsheets.
Enter the number of columns to compare.
Enter the names of the columns to compare in the first and second spreadsheets as prompted.

**Review the Output:**

The program will print any discrepancies found during the comparison.
If no discrepancies are found, the program will print "All information matches."
**Example**
Enter the path for the first spreadsheet (e.g., spreadsheets/spreadsheet1.xlsx): spreadsheets/spreadsheet1.xlsx
Enter the path for the second spreadsheet (e.g., spreadsheets/spreadsheet2.xlsx): spreadsheets/spreadsheet2.xlsx
Enter the name of the column for employee IDs in the first spreadsheet: ID
Enter the name of the column for employee IDs in the second spreadsheet: ID
Are the names of the columns the same on both spreadsheets? (yes/no): yes
Enter the number of columns to compare: 3
Enter the name of column 1 to compare: Salary
Enter the name of column 2 to compare: Bonus
Enter the name of column 3 to compare: Department
Please ignore the StatusLogger Error.
Mismatch found for ID 123:
  Salary vs Salary: 50000 vs 60000
  Bonus vs Bonus: 5000 vs 5500
  Department vs Department: HR vs Finance
All information matches.

**License**
This project is licensed under the MIT License. See the LICENSE file for more details.

**Contributing**
1. Fork the repository.
2. Create a new branch (git checkout -b feature-branch).
3. Make your changes.
4. Commit your changes (git commit -m 'Add new feature').
5. Push to the branch (git push origin feature-branch).
6. Open a pull request.

**Contact**
For any inquiries or issues, please contact me at liamkelso02@gmail.com.
