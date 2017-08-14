package com.innovative.excelfilereader;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.text.SimpleDateFormat;
import java.util.*;

class HelperMethods {

    //Method to loop over the rows in worksheet
    static void loopForRow(Iterator<Row> iterator, int noOfCol){
        int count = 0; // for Row counts

        while (iterator.hasNext()) {   // Loop for Row
            if (count < 2) printDashLine(noOfCol);
            count++;
            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            loopForColumn(cellIterator);
            System.out.println();
        }
    }

    //Method to loop over the columns in worksheet
    static void loopForColumn(Iterator<Cell> cellIterator){
        while (cellIterator.hasNext()) {  // Loop for Column
            Cell currentCell = cellIterator.next();
            printCell(currentCell);
        }
        System.out.println();

    }

    //Method for Printing Cells according to their type
    static void printCell(Cell currentCell){
        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            System.out.format("%-35s",wraptext(currentCell.getStringCellValue()));

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

    static void printUsage(){
        System.out.println("\nUsage: \nmvn exec:java -Dexec.args=\"FILEPATH [worksheet index]\" ");
        System.out.println("OR");
        System.out.println("java -jar FILEPATH [worksheet index]");
    }


    static void printDashLine(int columnCounts){
        for(int i=0; i<(columnCounts*40)-4; i++){
            System.out.print("-");
        }
        System.out.println();
    }

    static String wraptext(String str){
        final int FIXED_WIDTH = 30;
        String temp = "";
        if(str !=null && str.length() > FIXED_WIDTH) {
            temp = str.substring(0, FIXED_WIDTH) + "...";
        } else {
            temp = str;
        }
        return temp;
    }

    static void printSheetNames(Workbook workbook){
        for (int i=0; i<workbook.getNumberOfSheets(); i++) {
            System.out.println(i+1 + ". "+ workbook.getSheetName(i));
        }
        System.out.println();
    }

    static void printRowsColumnCountsOfWorksheet(Workbook workbook,int sheetIndex){
        Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
        int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();
        int noOfRows = datatypeSheet.getPhysicalNumberOfRows();
        String worksheetName = workbook.getSheetName(sheetIndex).toUpperCase();

        System.out.println(worksheetName + ": ");
        System.out.println("Row Counts: " + noOfRows);
        System.out.println("Column Counts: " + noOfCol);
    }

    static void printDataTypes(Workbook workbook, int sheetIndex){
        Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
        Row row = datatypeSheet.getRow(1);
        int count = 1;
        Iterator<Cell> cellIterator = row.iterator();

        System.out.print("Column DataTypes: ");
        while (cellIterator.hasNext()) {  // Loop for Column
            Cell currentCell = cellIterator.next();
            System.out.print("("+count+") "+getDataType(currentCell) + "\t");
            count++;
        }
        System.out.println("\n");
    }

    static String getDataType(Cell currentCell){
        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            return String.class.getSimpleName();

        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(currentCell)) {  // Checking if NUMERIC type is Date or not
                return Date.class.getSimpleName();
            }else{
                if (currentCell.getNumericCellValue() == Math.ceil(currentCell.getNumericCellValue())){ // checking if the value is float
                    return int.class.getSimpleName();
                }
                else{
                    return float.class.getSimpleName();
                }

            }
        }
        else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
            if(currentCell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                if (HSSFDateUtil.isCellDateFormatted(currentCell)) {
                    return Date.class.getSimpleName();
                }else{
                    return float.class.getSimpleName();
                }
            } else if(currentCell.getCachedFormulaResultTypeEnum() == CellType.STRING){
                return String.class.getSimpleName();
            }
            //   System.out.format("%-100s",currentCell.getCellFormula());
        }
        else if (currentCell.getCellTypeEnum() == CellType.BOOLEAN) {
            return boolean.class.getSimpleName();
        }
        return "";
    }
}
