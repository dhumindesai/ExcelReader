package com.innovative.excelfilereader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;



public class ApachePOIExcelReader {

    
    /*
        *This method displays the worksheet in table format and the summary.
        * Input: file path, worksheet number
        * Output: prints Table, Row Counts, Column Counts, Data Types of each column
     */
    public static void readWorksheet(String filePath, int sheetIndex){

        try {
            sheetIndex = sheetIndex - 1;

            // Getting Workbook objects and Row Iterator
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> iterator = datatypeSheet.iterator();
            int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();

            //printing table and summary
            HelperMethods.loopForRow(iterator,noOfCol);
            HelperMethods.printDashLine(noOfCol);
            HelperMethods.printRowsColumnCountsOfWorksheet(workbook,sheetIndex);
            HelperMethods.printDataTypes(workbook,sheetIndex);

        } catch (FileNotFoundException e) {
            System.out.println("File Path is invalid.");
        } catch (IllegalArgumentException e) {
            System.out.println("One of the arguments is incorrect.");
            HelperMethods.printUsage();
        } catch (IndexOutOfBoundsException e){
            System.out.println("No arguments are given.\n");
            HelperMethods.printUsage();
        } catch (IOException e){
            System.out.println("Could not load the excel file");
        }
    }

    /*
        * This method prints Metadata of the Excel File
        * Input:
        * @param filePath : it takes file path of the Excel File
        * Output: Prints Total No of worksheets, List of worksheets, summary of all worksheets
     */
    public static void printMetadata(String filePath){

        HelperMethods.printDashLine(2);
        System.out.println("\t\t\t\t\t\t\t\tMETADATA");
        HelperMethods.printDashLine(2);
        try{
            //get all worksheets in excel file
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new HSSFWorkbook(excelFile);
            int noOfWorksheets =  workbook.getNumberOfSheets();

            // printing Metadata
            System.out.println("Total Worksheets in excel file: "+noOfWorksheets + "\n");
            System.out.println("Worksheets: ");
            HelperMethods.printSheetNames(workbook);
            for(int i = 0; i < noOfWorksheets; i++) { //get Rows and Columns of each Worksheets
                HelperMethods.printRowsColumnCountsOfWorksheet(workbook, i);
                HelperMethods.printDataTypes(workbook, i);
            }
        }
        catch (FileNotFoundException e) {
            System.out.println("File Path is invalid.");
        } catch (IllegalArgumentException e) {
            System.out.println("One of the arguments is incorrect.");
            HelperMethods.printUsage();
        } catch (IndexOutOfBoundsException e){
            System.out.println("No arguments are given.\n");
            HelperMethods.printUsage();
        } catch (IOException e){
            System.out.println("Could not load the excel file");
        }

    }


    /*
        * This method filters the given worksheet by the column and condition and prints it.
        * input:
        * @param filePath : Path of Excel file
        * @param sheetIndex: Sheet number
        * @param columnNum: column number in the worksheet
        * @operator: '=' or '<' or '>' (for String DataType, only '=')
        * @operand: any number or string
        * output: prints filtered table
     */
    public static void printQuery(String filePath, int sheetIndex, int columnNum, char operator, String Operand){

        try {
            sheetIndex = sheetIndex - 1;
            columnNum = columnNum - 1;
            // Getting Workbook objects and Row Iterator
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> iterator = datatypeSheet.iterator();

            // Getting Worksheet information from Worksheet's object
            int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();
            int count = 0; // for Row counts
            int resultCount = 0;

            while (iterator.hasNext()) {   // Loop for Row

                Row currentRow = iterator.next();
                Cell currentCell = CellUtil.getCell(currentRow, columnNum);
                Iterator<Cell> cellIterator = currentRow.iterator();

                if (count < 1)
                {
                    HelperMethods.printDashLine(noOfCol);
                    HelperMethods.loopForColumn(cellIterator);
                    HelperMethods.printDashLine(noOfCol);
                }
                else {
                    try{
                        double num = Double.parseDouble(Operand);
                        // is an integer!
                        switch (operator){
                            case '=': if(currentCell.getNumericCellValue() == num){ HelperMethods.loopForColumn(cellIterator);resultCount++;}
                                break;
                            case '<': if(currentCell.getNumericCellValue() < num) { HelperMethods.loopForColumn(cellIterator);resultCount++;}
                                break;
                            case '>': if(currentCell.getNumericCellValue() > num) { HelperMethods.loopForColumn(cellIterator);resultCount++;}
                                break;
                            default: throw new IllegalArgumentException();
                        }

                        } catch (NumberFormatException e) {
                        // not an integer!
                        switch (operator){
                            case '=': if(currentCell.getStringCellValue().equals(Operand)){ HelperMethods.loopForColumn(cellIterator);resultCount++;}
                                break;
                            default: throw new IllegalArgumentException();
                        }
                    }

                }
                count++;
            }

            HelperMethods.printDashLine(noOfCol);
            System.out.println("Row Counts: "+resultCount);
            HelperMethods.printDataTypes(workbook,sheetIndex);


        } catch (FileNotFoundException e) {
            System.out.println("File Path is invalid.");
        } catch (IllegalArgumentException e) {
           // e.printStackTrace();
            System.out.println("One of the arguments is incorrect.");
            HelperMethods.printUsage();
        } catch (IndexOutOfBoundsException e){
            System.out.println("No arguments are given.\n");
            HelperMethods.printUsage();
        } catch (IOException e){
            System.out.println("Could not load the excel file");
        }
    }




}
