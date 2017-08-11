package com.innovative.excelfilereader;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.text.SimpleDateFormat;
import java.util.*;
public class HelperMethods {
    public static void printUsage(){
        System.out.println("\nUsage: \nmvn exec:java -Dexec.args=\"FILEPATH [worksheet index]\" ");
        System.out.println("OR");
        System.out.println("java -jar FILEPATH [worksheet index]");
    }


    public static void printDashLine(int columnCounts){
        for(int i=0; i<(columnCounts*40)-4; i++){
            System.out.print("-");
        }
        System.out.println();
    }

    public static String wraptext(String str){
        final int FIXED_WIDTH = 30;
        String temp = "";
        if(str !=null && str.length() > FIXED_WIDTH) {
            temp = str.substring(0, FIXED_WIDTH) + "...";
        } else {
            temp = str;
        }
        return temp;
    }

    public static void printSheetNames(Workbook workbook){
        for (int i=0; i<workbook.getNumberOfSheets(); i++) {
            System.out.println(i+1 + ". "+ workbook.getSheetName(i));
        }
        System.out.println();
    }

    public static void printRowsColumnCountsOfWorksheet(Workbook workbook,int sheetIndex){
        Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
        int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();
        int noOfRows = datatypeSheet.getPhysicalNumberOfRows();
        String worksheetName = workbook.getSheetName(sheetIndex).toUpperCase();

        System.out.println(worksheetName + ": ");
        System.out.println("Row Counts: " + noOfRows);
        System.out.println("Column Counts: " + noOfCol);
    }

    public static void printDataTypes(Workbook workbook, int sheetIndex){
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

    public static String getDataType(Cell currentCell){
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
