import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;
import java.util.Iterator;
import java.text.SimpleDateFormat;


public class ApachePOIExcelReader {

    public static void readWorksheet(String filePath, int sheetIndex){
        try {

            // Getting Workbook objects and Row Iterator
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> iterator = datatypeSheet.iterator();

            // Getting Worksheet information from Worksheet's object
            int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();
            int count = 0; // for Row counts

            while (iterator.hasNext()) {   // Loop for Row
                if (count < 2) HelperMethods.printDashLine(noOfCol);
                count++;
                int cellNum = 0; //  cell number  in the current row
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {  // Loop for Column
                    cellNum++;
                    Cell currentCell = cellIterator.next();
                    printCell(currentCell);
                }
                System.out.println();
            }
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
            System.out.println("Total Worksheets in excel file: "+noOfWorksheets + "\n");
            System.out.println("Worksheets: ");
            HelperMethods.printSheetNames(workbook);

            //get Rows and Columns of each Worksheets
            for(int i = 0; i < noOfWorksheets; i++) {
                Sheet datatypeSheet = workbook.getSheetAt(i);
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



}
