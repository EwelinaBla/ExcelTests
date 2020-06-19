package Objects.CreateFile;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


public class CreateFile {
    HSSFWorkbook workbook = new HSSFWorkbook ();
    HSSFSheet sheet = workbook.createSheet ("Static");

    public void createFileExcel(String[] columnB) {
        setHeaders ();
        setInputStaticData (columnB);
        setForceData ();
        outFile ();
    }

    private void setHeaders() {
        String[] header = new String[]{"Dane wejściowe", "Wartość", "Jednostka"};
        HSSFRow rowHeader = sheet.createRow (0);
        for (int column = 0; column < header.length; column++) {
            Cell cell = rowHeader.createCell (column);
            cell.setCellValue (header[column]);
            cell.setCellStyle (setHeaderStyle ());
        }

        String[] headerSecond = new String[]{"Oznaczenie", "Wartość", "Jednostka"};
        HSSFRow rowHeaderSecond = sheet.createRow (8);
        for (int column = 0; column < headerSecond.length; column++) {
            Cell cell = rowHeaderSecond.createCell (column);
            cell.setCellValue (headerSecond[column]);
            cell.setCellStyle (setHeaderStyle ());
        }
    }

    private void setInputStaticData(String[] columnB) {

        String[] columnA = new String[]{"P1", "a", "b", "D1", "D2", "Beta"};
        for (int i = 0; i < columnA.length; i++) {
            Cell cell = sheet.createRow (i + 1).createCell (0);
            cell.setCellValue (columnA[i]);
            cell.setCellStyle (setDefaultStyle ());
        }

        String[] columnC = new String[]{"N", "m", "m", "m", "m", "rad"};
        for (int i = 0; i < columnC.length; i++) {
            Cell cell = sheet.createRow (i + 1).createCell (2);
            cell.setCellValue (columnC[i]);
            cell.setCellStyle (setDefaultStyle ());
        }

        try {
            for (int i = 0; i < columnB.length; i++) {
                Cell cell = sheet.getRow (i + 1).createCell (1);
                cell.setCellValue (Double.parseDouble (columnB[i]));
                cell.setCellStyle (setDefaultStyle ());
            }
        } catch (NumberFormatException e) {

        }
    }

    private void setForceData() {

        String[] force = new String[]{"P1y", "P1z", "P2y", "P2z"};
        String[] valueForce = new String[]{"ROUND((B2*COS(B7)),2)", "ROUND((B2*SIN(B7)),2)", "ROUND((B13*TAN(B7)),2)", "ROUND((B10*(B5/B6)),2)"};

        for (int i = 0; i < force.length; i++) {
            HSSFRow row = sheet.createRow (i + 9);
            Cell cellColumnA = row.createCell (0);
            cellColumnA.setCellValue (force[i]);
            cellColumnA.setCellStyle (setDefaultStyle ());

            Cell cellColumnB = row.createCell (1);
            cellColumnB.setCellFormula (valueForce[i]);
            cellColumnB.setCellStyle (setDefaultStyle ());

            Cell cellColumnC = row.createCell (2);
            cellColumnC.setCellValue ("N");
            cellColumnC.setCellStyle (setDefaultStyle ());
        }
    }

    private CellStyle setHeaderStyle() {
        HSSFFont font = workbook.createFont ();
        font.setBoldweight (HSSFFont.BOLDWEIGHT_BOLD);
        font.setItalic (true);
        font.setFontName ("Arial");
        font.setColor (Font.COLOR_RED);

        HSSFCellStyle style = workbook.createCellStyle ();
        borderStyle (style);
        style.setFillForegroundColor (HSSFColor.DARK_BLUE.index);
        style.setFillPattern ((CellStyle.SOLID_FOREGROUND));
        style.setAlignment (CellStyle.ALIGN_CENTER);

        style.setFont (font);
        return style;
    }

    private CellStyle setDefaultStyle() {
        HSSFCellStyle defaultStyle = workbook.createCellStyle ();
        borderStyle (defaultStyle);
        sheet.setDefaultColumnWidth (15);
        return defaultStyle;
    }

    private void borderStyle(CellStyle cellStyle) {
        cellStyle.setBorderBottom (CellStyle.BORDER_THIN);
        cellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
        cellStyle.setBorderTop (CellStyle.BORDER_THIN);
        cellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());
        cellStyle.setBorderLeft (CellStyle.BORDER_THIN);
        cellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
        cellStyle.setBorderRight (CellStyle.BORDER_THIN);
        cellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());

        cellStyle.setWrapText (true);
    }


    private void outFile() {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream (new File ("C:\\Projects\\ProjectExcelFile\\calculator.xls"));
            workbook.write (fileOutputStream);
            fileOutputStream.close ();
        } catch (FileNotFoundException e) {
            e.printStackTrace ();
        } catch (IOException e) {
            e.printStackTrace ();
        }
    }


}


//        HSSFRow rowP1y = sheet.createRow (9);
//        rowP1y.createCell (0).setCellValue ("P1y");
//        rowP1y.createCell (1).setCellFormula ("B2*COS(B7)");
//        rowP1y.createCell (2).setCellValue ("N");
//
//
//        sheet.createRow (10).createCell (0).setCellValue ("P1z");
//        sheet.createRow (10).createCell (1).setCellFormula ("B2*SIN(B7)");
//        sheet.createRow (10).createCell (2).setCellValue ("N");
//
//        sheet.createRow (11).createCell (0).setCellValue ("P2y");
//        sheet.createRow (11).createCell (1).setCellFormula ("B13*TAN(B7)");
//        sheet.createRow (11).createCell (2).setCellValue ("N");
//
//        sheet.createRow (12).createCell (0).setCellValue ("P1z");
//        sheet.createRow (12).createCell (1).setCellFormula ("B10*(B5/B6)");
//        sheet.createRow (12).createCell (2).setCellValue ("N");


//        HashMap<Integer, String[]> inputData = new HashMap<> ();
//        inputData.put (0, new String[]{"Dane wejściowe", "Wartość", "Jednostka"});
//        inputData.put (1, new String[]{"P1", , "N"});
//        inputData.put (2, new String[]{"a", valueA, "m"});
//        inputData.put (3, new String[]{"b", valueB, "m"});
//        inputData.put (4, new String[]{"D1", valueD1, "m"});
//        inputData.put (5, new String[]{"D2", valueD2, "m"});
//        inputData.put (6, new String[]{"Beta", valueBeta, "rad"});
//        inputData.put (8, new String[]{"Oznaczenie", "Wartość", "Jednostka"});
//        inputData.put (9, new String[]{"P1y", String.valueOf (valueP1y), "N"});
//            inputData.put (10, new String[]{"P1z", String.valueOf (valueP1z), "N"});
//            inputData.put (11, new String[]{"P2y", String.valueOf (valueP2y), "N"});
//            inputData.put (12, new String[]{"P2z", String.valueOf (valueP2z), "N"});
//
//        Set<Integer> keysInputData = inputData.keySet ();

//        for (Integer key : keysInputData) {
//            HSSFRow row = sheet.createRow (key);
//            String[] valuesKeys = inputData.get (key);
//            int cellum = 0;
//            for (String valueKey : valuesKeys) {
//                Cell cell = row.createCell (cellum++);
//                cell.setCellValue (valueKey);
//            }

