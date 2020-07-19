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


public class CreateFileExcelCalculator {
    private HSSFWorkbook workbook = new HSSFWorkbook ();
    private HSSFSheet sheet = workbook.createSheet ("Static");

    public void createFileExcel(String[] columnB, String path) {
        setHeaders ();
        setInputStaticData (columnB);
        setForcesData ();
        saveToFile (path);
    }

    private void setHeaders() {
        String[] firstHeader = new String[]{"Dane wejściowe", "Wartość", "Jednostka"};
        HSSFRow rowHeader = sheet.createRow (0);
        for (int column = 0; column < firstHeader.length; column++) {
            Cell cell = rowHeader.createCell (column);
            cell.setCellValue (firstHeader[column]);
            cell.setCellStyle (setHeaderStyle ());
        }

        String[] secondHeader = new String[]{"Oznaczenie", "Wartość", "Jednostka"};
        HSSFRow rowHeaderSecond = sheet.createRow (8);
        for (int column = 0; column < secondHeader.length; column++) {
            Cell cell = rowHeaderSecond.createCell (column);
            cell.setCellValue (secondHeader[column]);
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
            Cell cell = sheet.getRow (i + 1).createCell (2);
            cell.setCellValue (columnC[i]);
            cell.setCellStyle (setDefaultStyle ());
        }

        for (int i = 0; i < columnB.length; i++) {
            Cell cell = sheet.getRow (i + 1).createCell (1);
            cell.setCellValue (Double.parseDouble (columnB[i]));
            cell.setCellStyle (setDefaultStyle ());
            }
    }

    private void setForcesData() {

        String[] forces = new String[]{"P1y", "P1z", "P2y", "P2z"};
        String[] valuesForces = new String[]{"ROUND((B2*COS(B7)),2)", "ROUND((B2*SIN(B7)),2)", "ROUND((B13*TAN(B7)),2)", "ROUND((B10*(B5/B6)),2)"};

        for (int i = 0; i < forces.length; i++) {
            HSSFRow row = sheet.createRow (i + 9);
            Cell cellColumnA = row.createCell (0);
            cellColumnA.setCellValue (forces[i]);
            cellColumnA.setCellStyle (setDefaultStyle ());

            Cell cellColumnB = row.createCell (1);
            cellColumnB.setCellFormula (valuesForces[i]);
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
        setBorderStyle (style);
        style.setFillForegroundColor (HSSFColor.DARK_BLUE.index);
        style.setFillPattern ((CellStyle.SOLID_FOREGROUND));
        style.setAlignment (CellStyle.ALIGN_CENTER);

        style.setFont (font);
        return style;
    }

    private CellStyle setDefaultStyle() {
        HSSFCellStyle defaultStyle = workbook.createCellStyle ();
        setBorderStyle (defaultStyle);
        sheet.setDefaultColumnWidth (15);
        return defaultStyle;
    }

    private void setBorderStyle(CellStyle cellStyle) {
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

    private void saveToFile(String path) {
        try(FileOutputStream fileOutputStream = new FileOutputStream (new File (path))) {
            workbook.write (fileOutputStream);
        } catch (FileNotFoundException e) {
            System.out.println (e.getMessage ());
        } catch (IOException e) {
            System.out.println ("Błąd zapisu");
            System.exit (1);
        }
    }
}



