import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;
import java.text.SimpleDateFormat;


public class ApachePOIExcelReader {

    public static void readExcel(String FILE_NAME, int sheetIndex){
        try {

            // Getting Workbook objects and Row Iterator
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> iterator = datatypeSheet.iterator();

            // Getting Worksheet information from Worksheet's object
            int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();
            int noOfRows = datatypeSheet.getPhysicalNumberOfRows();
            String worksheetName = datatypeSheet.getSheetName().toUpperCase();
            int count = 0; // for Row counts


            while (iterator.hasNext()) {   // Loop for Row
                if (count < 2) printDashLine(noOfCol);
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
            printDashLine(noOfCol);

            //print Total Rows and Columns in the worksheet
            System.out.println("Current Worksheet: "+worksheetName);
            System.out.println("Total Rows: "+noOfRows);
            System.out.println("Total Columns: "+noOfCol);

        } catch (FileNotFoundException e) {
            System.out.println("File Path is invalid. Please provide the valid file path.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    //Method for Printing Cells according to their type
    public static void printCell(Cell currentCell){
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

    }

    public static void printDashLine(int columnCounts){
        for(int i=0; i<columnCounts*35; i++){
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

}
