package com.jaxrxapachepoi.rest.utils;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Date;
import java.util.Locale;

public class ExcelUtils {
//https://stackoverflow.com/a/44969043
    public static void compareExcelData(File f1, File f2, String if1Sheet, String if2Sheet, int[] columns){
        try{
            OPCPackage f1opcPackage = OPCPackage.open(f1);
            XSSFWorkbook f1workbook = new XSSFWorkbook(f1opcPackage);

            OPCPackage f2opcPackage = OPCPackage.open(f2);
            XSSFWorkbook f2workbook = new XSSFWorkbook(f2opcPackage);

            XSSFSheet f1Sheet = f1workbook.getSheet("f1Sheet");
            XSSFSheet f2Sheet = f1workbook.getSheet("f2Sheet");

            SXSSFWorkbook finalResult = new SXSSFWorkbook(100);
            Sheet resultSheet = finalResult.createSheet();

            for (int j = 0; j < f1Sheet.getPhysicalNumberOfRows(); j++) {
                if (f2Sheet.getPhysicalNumberOfRows() <= j) return;

                XSSFRow f1row = f1Sheet.getRow(j);
                XSSFRow f2row = f2Sheet.getRow(j);

                Row resultRow = resultSheet.createRow(j);

                if ((f1row == null) || (f2row == null)) {
                    continue;
                }

                compareDataInRow(f1row, f2row,resultRow);
            }

        }catch (Exception e){

        }

    }

    private static void compareDataInRow(XSSFRow f1row, XSSFRow f2row, Row resultRow) {
        int cellCreater = 0;
        for (int k = 0; k < f1row.getLastCellNum(); k++) {
            if (f2row.getPhysicalNumberOfCells() <= k) return;
            Cell cell1 = resultRow.createCell(cellCreater);
            Cell cell2 = resultRow.createCell(cellCreater+1);
            Cell cell3 = resultRow.createCell(cellCreater+2);

            XSSFCell f1cell = f1row.getCell(k);
            XSSFCell f2cell = f2row.getCell(k);

            if ((f1cell == null) || (f2cell == null)) {
                continue;
            }

            //compareDataInCell(f1cell, f2cell);

            cell1.setCellValue(f1cell.getStringCellValue());
            cell2.setCellValue(f1cell.getStringCellValue());
            cell3.setCellValue(compareDataInCell(f1cell, f2cell) == true?"MATCHED":"NOT MATCHED");

            cellCreater += 3;
        }
    }

    private static boolean compareDataInCell(XSSFCell f1cell, XSSFCell f2cell) {
        boolean notmatched = false;
        if (isCellTypeMatches(f1cell, f2cell)) {
            final CellType loc1cellType = f1cell.getCellTypeEnum();
            switch(loc1cellType) {
                case BLANK:
                case STRING:
                case ERROR:
                    notmatched = isCellContentMatches(f1cell, f2cell);
                    break;
                case BOOLEAN:
                    notmatched = isCellContentMatchesForBoolean(f1cell, f2cell);
                    break;
                case FORMULA:
                    notmatched = isCellContentMatchesForFormula(f1cell, f2cell);
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(f1cell)) {
                        notmatched = isCellContentMatchesForDate(f1cell, f2cell);
                    } else {
                        notmatched = isCellContentMatchesForNumeric(f1cell, f2cell);
                    }
                    break;
                default:
                    throw new IllegalStateException("Unexpected cell type: " + loc1cellType);
            }

        }
        return notmatched;
    }

    private static boolean isCellTypeMatches(XSSFCell f1cell, XSSFCell f2cell) {
        CellType type1 = f1cell.getCellTypeEnum();
        CellType type2 = f2cell.getCellTypeEnum();
        if (type1 == type2) return true;
        //addMessage(loc1, loc2, "Cell Data-Type does not Match in :: ",type1.name(), type2.name());
        return false;
    }

    private static boolean isCellContentMatches(XSSFCell f1cell, XSSFCell f2cell) {
        // TODO: check for null and non-rich-text cells
        String str1 = f1cell.getRichStringCellValue().getString();
        String str2 = f2cell.getRichStringCellValue().getString();
        return !str1.equals(str2);
        /*if (!str1.equals(str2)) {
            return true;
            //addMessage(loc1,loc2,CELL_DATA_DOES_NOT_MATCH,str1,str2);
        }
        return false;*/
    }

    /**
     * Checks if cell content matches for boolean.
     */
    private static boolean isCellContentMatchesForBoolean(XSSFCell f1cell, XSSFCell f2cell) {
        boolean b1 = f1cell.getBooleanCellValue();
        boolean b2 = f2cell.getBooleanCellValue();
        return b1 != b2;
        /*if (b1 != b2) {
            //addMessage(loc1,loc2,CELL_DATA_DOES_NOT_MATCH,Boolean.toString(b1),Boolean.toString(b2));
        }*/
    }

    /**
     * Checks if cell content matches for date.
     */
    private static boolean isCellContentMatchesForDate(XSSFCell f1cell, XSSFCell f2cell) {
        Date date1 = f1cell.getDateCellValue();
        Date date2 = f2cell.getDateCellValue();
        return !date1.equals(date2);
        /*if (!date1.equals(date2)) {
            //addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, date1.toString(), date2.toString());
        }*/
    }


    /**
     * Checks if cell content matches for formula.
     */
    private static boolean isCellContentMatchesForFormula(XSSFCell f1cell, XSSFCell f2cell) {
        // TODO: actually evaluate the formula / NPE checks
        String form1 = f1cell.getCellFormula();
        String form2 = f2cell.getCellFormula();
        return !form1.equals(form2);
        /*if (!form1.equals(form2)) {
            //addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, form1, form2);
        }*/
    }

    /**
     * Checks if cell content matches for numeric.
     */
    private static boolean isCellContentMatchesForNumeric(XSSFCell f1cell, XSSFCell f2cell) {
        // TODO: Check for NaN
        double num1 = f1cell.getNumericCellValue();
        double num2 = f2cell.getNumericCellValue();
        return num1 != num2;
        /*if (num1 != num2) {
            //addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Double.toString(num1), Double.toString(num2));
        }*/
    }
    private void addMessage(XSSFWorkbook loc1, XSSFWorkbook loc2, String messageStart, String value1, String value2) {
        /*String str =
                String.format(Locale.ROOT, "%s\nworkbook1 -> %s -> %s [%s] != workbook2 -> %s -> %s [%s]",
                        messageStart,
                        loc1.sheet.getSheetName(), new CellReference(loc1.cell).formatAsString(), value1,
                        loc2.sheet.getSheetName(), new CellReference(loc2.cell).formatAsString(), value2
                );
        listOfDifferences.add(str);*/
    }
}
