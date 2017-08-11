package com.innovative;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.util.CellUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.text.SimpleDateFormat;


public class ApachePOIExcelReader {

    // takes file path and worksheet number and print the contents and metadata of given worksheet
    public static void readWorksheet(String filePath, int sheetIndex){
        try {
            sheetIndex = sheetIndex - 1;

            // Getting Workbook objects and Row Iterator
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> iterator = datatypeSheet.iterator();
            int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();

            loopForRow(iterator,noOfCol);
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

    //Method for MetaData
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
               // Sheet datatypeSheet = workbook.getSheetAt(i);
                HelperMethods.printRowsColumnCountsOfWorksheet(workbook, i);
                HelperMethods.printDataTypes(workbook, i);
            }
        }
        catch(FileNotFoundException e){
            System.out.println("The File Path is Invalid.");

        }catch(IOException e){
            System.out.println("Could not load Excel File.");
        }

    }




    public static void printQuery(String filePath, int sheetIndex, int columnNum, char operator, String Operand){
        try {
            sheetIndex = sheetIndex - 1;
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
                    loopForColumn(cellIterator);
                    HelperMethods.printDashLine(noOfCol);
                }
                else {
                    try{
                        double num = Double.parseDouble(Operand);
                        // is an integer!
                        switch (operator){
                            case '=': if(currentCell.getNumericCellValue() == num){ loopForColumn(cellIterator);resultCount++;}
                                break;
                            case '<': if(currentCell.getNumericCellValue() < num) { loopForColumn(cellIterator);resultCount++;}
                                break;
                            case '>': if(currentCell.getNumericCellValue() > num) { loopForColumn(cellIterator);resultCount++;}
                                break;
                            default: throw new IllegalArgumentException();
                        }

                        } catch (NumberFormatException e) {
                        // not an integer!
                        switch (operator){
                            case '=': if(currentCell.getStringCellValue().equals(Operand)){ loopForColumn(cellIterator);resultCount++;}
                                break;
                            default: throw new IllegalArgumentException();
                        }
                    }

                }
                count++;
               // System.out.println();
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


    //Method for Printing Cells according to their type
    public static void printCell(Cell currentCell){
        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            System.out.format("%-35s",HelperMethods.wraptext(currentCell.getStringCellValue()));

        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(currentCell)) {  // Checking if NUMERIC type is Date or not
                System.out.format("%-35s",new SimpleDateFormat("MM/dd/yyyy").format(currentCell.getDateCellValue()));
            }else{
                if (currentCell.getNumericCellValue() == Math.ceil(currentCell.getNumericCellValue())){ // checking if the value is float
                    System.out.format("%-35d",(int)currentCell.getNumericCellValue());
                }
                else{
                    System.out.format("%-35.1f",currentCell.getNumericCellValue());
                }

            }
        }
        else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
            if(currentCell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                if (HSSFDateUtil.isCellDateFormatted(currentCell)) {
                    System.out.format("%-35s",new SimpleDateFormat("MM/dd/yyyy").format(currentCell.getDateCellValue()));
                }else{
                    System.out.format("%-35.1f",currentCell.getNumericCellValue());
                }
            } else if(currentCell.getCachedFormulaResultTypeEnum() == CellType.STRING){
                System.out.format("%-35s",currentCell.getStringCellValue());
            }
            //   System.out.format("%-100s",currentCell.getCellFormula());
        }
        else if (currentCell.getCellTypeEnum() == CellType.BOOLEAN) {
            System.out.format("%-35s",currentCell.getBooleanCellValue());
        }
        else if (currentCell.getCellTypeEnum() == CellType.BLANK) {
            System.out.format("%-35s","NULL");
        }

        System.out.format("%-5s","|");

    }

    public static void loopForRow(Iterator<Row> iterator, int noOfCol){
        int count = 0; // for Row counts

        while (iterator.hasNext()) {   // Loop for Row
            if (count < 2) HelperMethods.printDashLine(noOfCol);
            count++;
            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            loopForColumn(cellIterator);
            System.out.println();
        }
    }

    public static void loopForColumn(Iterator<Cell> cellIterator){
        while (cellIterator.hasNext()) {  // Loop for Column
            //cellNum++;
            Cell currentCell = cellIterator.next();
            printCell(currentCell);
        }
        System.out.println();
    }




}
