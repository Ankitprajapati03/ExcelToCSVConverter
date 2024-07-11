package com.example.demo;

//importing required packages
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

// created a Class for Conversion Excel to Csv

public class ExcelToCsvConversion {
    public static void main(String args[]) {
        // Path to the configuration Excel sheet
        String configurSheetPath = "D://Excel to CSV//CSD_TO_CSV.xlsx";
        // Path to the data Excel sheet
        String dataSheetPath = "D://Excel to CSV//CSD - Internal.xlsx";

        // Call the configureSheet method to process the sheet
        configureSheet(configurSheetPath, dataSheetPath);
    }

    /**
     * Configures processing of a sheet based on configuration data from another
     * Excel sheet.
     * 
     * @param configurSheetPath Path to the configuration Excel sheet
     * @param dataSheetPath     Path to the Data excel Sheet
     * @param sheetName         the name of the current sheet being processed
     * @param sheet             the current sheet object being processed
     */
    public static void configureSheet(String configurSheetPath, String dataSheetPath) {
        try {
            // Read configuration Excel sheet
            FileInputStream configInputStream = new FileInputStream(new File(configurSheetPath));
            Workbook ConfigurWorkbook = WorkbookFactory.create(configInputStream);
            Sheet ConfigurSheet = ConfigurWorkbook.getSheetAt(0);

            // Iterate through each row of the configuration sheet
            for (int r = 1; r <= ConfigurSheet.getLastRowNum(); r++) {
                Row configurRow = ConfigurSheet.getRow(r);

                // Get sheet name from configuration
                Cell configurCell = configurRow.getCell(1);
                String configurSheetName = configurCell.getStringCellValue();

                // Get directory path from configuration
                Cell dirCell = configurRow.getCell(2);
                String dirname = dirCell.getStringCellValue();
                File file = new File(dirname);
                String dirPath = file.getParent(); // Extracts the directory path
                String csvName = file.getName();

                // Create directory for CSV files based on configuration
                String dirPathLocation = createDirectory("D://Excel to CSV//", dirPath, csvName);

                // Get transpose flag from configuration
                Cell transposeCell = configurRow.getCell(3);
                boolean IsTranspose = transposeCell.getBooleanCellValue();

                // Get comment flag from configuration
                Cell commentCell = configurRow.getCell(4);
                boolean IsComment = commentCell.getBooleanCellValue();

                Cell rangeCell = configurRow.getCell(5);
                String range = rangeCell.getStringCellValue();

                int colInd = 1;
                int rowInd = IsTranspose ? 2 : 0;

                // Read data from data Excel sheet based on configured indices
                List<List<String>> excelData = readExcel(dataSheetPath, configurSheetName, rowInd, colInd, IsComment);
                // System.out.println(excelData);

                // Transpose data if needed based on configuration
                if (IsTranspose && range.equals("na")) {
                    excelData = transposeData(excelData);
                }

                if (!range.equals("na") && IsTranspose) {

                    excelData = specificRange(excelData, range);

                    excelData = transposeData(excelData);

                }
                // if(!range.equals("na") && !IsTranspose)
                // {
                // excelData=specificRange(excelData,range);

                // }
                writeCsv(dirPathLocation, excelData);

            }

        } catch (Exception e) {
            System.out.println(e);
        }
    }

    /**
     * Reads data from an Excel sheet and returns it as a list of lists of strings.
     *
     * @param filePath  the path to the Excel file
     * @param sheet     the sheet to read from excel file
     * @param rowInd    the starting row index to read from sheet
     * @param colInd    the starting column index to read from row
     * @param IsComment Flag indicating whether to ignore comment cells.
     * @return a list of lists of strings containing the data from the Excel sheet
     */
    public static List<List<String>> readExcel(String dataSheetPath, String configurSheetName, int rowInd, int colInd,
            boolean IsComment) {
        List<List<String>> data = new ArrayList<>();
        try {
            // Create a File object for the data Excel sheet
            File datafile = new File(dataSheetPath);
            // Create a FileInputStream to read the Excel file
            FileInputStream fileIntput = new FileInputStream(datafile);
            // Create a Workbook object to access the Excel file's contents
            try (Workbook workbook = new XSSFWorkbook(fileIntput)) {
               
                // Iterate over each sheet in the workbook
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    // Get the current sheet
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();
                    if (sheetName.equals(configurSheetName)) {

                        // Iterate through each row of the sheet starting from the specified row index
                        for (int ro = rowInd; ro <= sheet.getLastRowNum(); ro++) {
                            Row row = sheet.getRow(ro);

                            // Create a list to store data for a single row
                            List<String> rowData = new ArrayList<>();

                            if (row != null) {
                                // Iterate through each cell of the row starting from the specified cell index
                                for (int j = colInd; j < sheet.getRow(0).getLastCellNum(); j++) {

                                    Cell cell = row.getCell(j);

                                    // skip the comment column if needed based on configuration
                                    if (!IsComment && j == sheet.getRow(0).getLastCellNum() - 1) {
                                        continue;
                                    }

                                    if (cell != null) {

                                        rowData.add(cellToString(cell));
                                    }
                                }
                            }

                            data.add(rowData);
                        }

                    }
                }
            }

        } catch (Exception exception) {
            System.out.println(exception);
        }
        return data;
    }

    /**
     * Converts a cell's value to a string.
     * 
     * @return the string representation of the cell's value
     *         this method is return cell value if value is any type like--
     *         String,Numeric,Boolean,Even formula and blank
     */
    private static String cellToString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numericValue = cell.getNumericCellValue();
                    // Check if the numeric value is an integer
                    if (numericValue == (long) numericValue) {
                        return Long.toString((long) numericValue);
                    } else {
                        return Double.toString(numericValue);
                    }
                }

            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "Unknown Cell Type";
        }
    }

    /**
     * Writes the Excel data to a CSV file.
     * 
     * @param csvFilePath the path to the CSV file to be created
     * @param excelData   the data from the Excel sheet to be written to the CSV
     *                    file
     */
    public static void writeCsv(String csvFilePath, List<List<String>> excelData) {

        try (FileWriter csvWriter = new FileWriter(csvFilePath)) {
            writeCsvWithModifiedHeader(csvWriter, excelData);

            // Iterate through each row of Excel data
            for (int i = 1; i < excelData.size(); i++) {
                List<String> row = excelData.get(i);
                for (int j = 0; j < row.size(); j++) {
                    csvWriter.append(row.get(j) != null ? escapeSpecialCharacters(row.get(j)) : "");
                    if (j < row.size() - 1) {
                        csvWriter.append(",");
                    }
                }
                csvWriter.append("\n");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void writeCsvWithModifiedHeader(FileWriter csvWriter, List<List<String>> data) {

        // Process the first row
        try {
            List<String> headerRow = data.get(0);
            for (int i = 0; i < headerRow.size(); i++) {
                String header = headerRow.get(i);
                if (header != null) {
                    header = header.toLowerCase().replace(" ", "_");
                }
                csvWriter.append(header != null ? header : "");
                if (i < headerRow.size() - 1) {
                    csvWriter.append(",");
                }
            }
            csvWriter.append("\n");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Escapes special characters in a string for CSV format.
     *
     * @param column the string to be escaped
     * @return the escaped string
     */
    private static String escapeSpecialCharacters(String column) {
        if (column.contains(",") || column.contains("\n") || column.contains("\"")) {
            // If the column contains a comma, newline, or double quotes, surround it with
            // double quotes
            column = column.replace("\"", "\"\""); // Escape double quotes by doubling them
            return "\"" + column + "\"";
        } else {
            return column;
        }
    }

    /**
     * Transposes the provided Excel data.
     *
     * @param excelData the data from the Excel sheet to be transposed
     * @return the transposed data
     */
    public static List<List<String>> transposeData(List<List<String>> excelData) {
        int rows = excelData.size();
        int maxCols = 0;

        // Determine the maximum number of columns in the data
        for (List<String> row : excelData) {
            if (row.size() > maxCols) {
                maxCols = row.size();
            }
        }
        List<List<String>> transpose = new ArrayList<>();

        // Iterate through each column index up to maxCols - 1 For Now for ignore
        // comment iterate the each Column index up to maxCols - 1
        for (int j = 0; j < maxCols; j++) {
            List<String> row = new ArrayList<>();
            for (int i = 0; i < rows; i++) {

                if (j < excelData.get(i).size()) {
                    row.add(excelData.get(i).get(j));
                }
                // else {
                // row.add(""); // Add empty string if column doesn't exist in original
                // }
            }
            transpose.add(row);
        }
        return transpose;
    }

    public static List<List<String>> specificRange(List<List<String>> excelData, String range) {

        List<List<String>> specificList = new ArrayList<>();

        if (excelData.isEmpty() || range == null || range.isEmpty()) {
            return specificList;
        }

        if (range.contains(",")) {
            // Split the range string into individual ranges or indices
            String[] ranges = range.split(",");
            for (String r : ranges) {
                if (r.contains("-")) {
                    // Handle range
                    String[] bounds = r.split("-");

                    try {
                        int start = Integer.parseInt(bounds[0].trim());
                        int end = Integer.parseInt(bounds[1].trim());
                        for (int i = start; i <= end; i++) {
                            addRowIfValid(excelData, specificList, i);
                         
                        }
                    } catch (NumberFormatException | ArrayIndexOutOfBoundsException e) {
                        // Handle error (e.g., log it or print a message)
                        System.err.println("Invalid range format: " + r);
                    }
                } else {
                    // Handle single index
                    try {
                        int index = Integer.parseInt(r.trim());
                        addRowIfValid(excelData, specificList, index);
                     
                    } catch (NumberFormatException | ArrayIndexOutOfBoundsException e) {
                        // Handle error (e.g., log it or print a message)
                        System.err.println("Invalid index format: " + r);
                    }
                }
            }
        }
         // Check if range is a single value (number or string)
        else if (!range.contains(",") && !range.contains("-")) {
            try {
                // Try to parse range as an integer index
                int index = Integer.parseInt(range.trim());
                addRowIfValid(excelData, specificList, index);
            } catch (NumberFormatException e) {
                // Handle string range by matching the first column value
                System.err.println("Invalid index format: " + range);
            }
            return specificList;
        }

        return specificList;
    }

    private static void addRowIfValid(List<List<String>> excelData, List<List<String>> specificList, int index) {
        // Adjust index to 0-based
        index = index - 1;
        if (index >= 0 && index < excelData.size()) {
            List<String> row = excelData.get(index);
            if (row != null && !row.isEmpty()) {
                specificList.add(row);
            }
        }
    }

    /**
     * this method is created for make directories
     * 
     * @param outputDirectory the base output directory
     * @param dirPath         the directory path to be created within the output
     *                        directory
     * @param csvName         the name of the CSV file
     * @return the full file path including the CSV file name
     */
    private static String createDirectory(String outputDirectory, String dirPath, String csvName) {
        // Split the sheetPath to get the first part of the path
        String directoryPath = outputDirectory + File.separator + dirPath;

        // Create the directories if they do not exist
        File directory = new File(directoryPath);
        if (!directory.exists()) {
            directory.mkdirs();
        }

        // Return the full file path including the CSV file name
        return directoryPath + File.separator + csvName;
    }

}
