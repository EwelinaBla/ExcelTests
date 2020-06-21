package Tests;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;

public class FileExcelTest extends BaseTest {
    private HSSFSheet sheet;

    @BeforeAll
    @DisplayName("Inicjalizacja testów")
    private void init() {
        openFile ();
        this.sheet = workbook.getSheetAt (0);
    }

    @Test
    public void checkNameOfSheet() {
        String nameSheet = sheet.getSheetName ();

        Assertions.assertEquals ("Static", nameSheet);
    }

    @Test
    public void checkNumberOfSheets() {
        int numberSheets = workbook.getNumberOfSheets ();

        Assertions.assertEquals (1, numberSheets);
    }

    @Test
    public void checkInputData() {
        double valueP1 = sheet.getRow (1).getCell (1).getNumericCellValue ();
        double valueA = sheet.getRow (2).getCell (1).getNumericCellValue ();
        double valueB = sheet.getRow (3).getCell (1).getNumericCellValue ();
        double valueD1 = sheet.getRow (4).getCell (1).getNumericCellValue ();
        double valueD2 = sheet.getRow (5).getCell (1).getNumericCellValue ();
        double valueBeta = sheet.getRow (6).getCell (1).getNumericCellValue ();

        Assertions.assertAll (
                () -> Assertions.assertTrue (valueP1 > 0),
                () -> Assertions.assertTrue (valueA > 0),
                () -> Assertions.assertTrue (valueB > 0),
                () -> Assertions.assertTrue (valueD1 > 0),
                () -> Assertions.assertTrue (valueD2 > 0),
                () -> Assertions.assertTrue (valueBeta > 0)
        );
    }

    @Test
    public void checkFormula() {
        String valueP1y = sheet.getRow (9).getCell (1).getCellFormula ();
        String valueP1z = sheet.getRow (10).getCell (1).getCellFormula ();
        String valueP2y = sheet.getRow (11).getCell (1).getCellFormula ();
        String valueP2z = sheet.getRow (12).getCell (1).getCellFormula ();

        String expectedValueP1y = "ROUND((B2*COS(B7)),2)";
        String expectedValueP1z = "ROUND((B2*SIN(B7)),2)";
        String expectedValueP2y = "ROUND((B13*TAN(B7)),2)";
        String expectedValueP2z = "ROUND((B10*(B5/B6)),2)";

        Assertions.assertAll (
                () -> Assertions.assertEquals (expectedValueP1y, valueP1y),
                () -> Assertions.assertEquals (expectedValueP1z, valueP1z),
                () -> Assertions.assertEquals (expectedValueP2y, valueP2y),
                () -> Assertions.assertEquals (expectedValueP2z, valueP2z)
        );
    }

    @Test
    public void checkUnit() {
        String poundalForFirstTable = sheet.getRow (1).getCell (2).getStringCellValue ();
        String unitOfAngle = sheet.getRow (6).getCell (2).getStringCellValue ();

        ArrayList<String> poundalForSecondTable = new ArrayList<> ();
        for (int i = 0; i < 4; i++) {
            String poundalSecondTable = sheet.getRow (i + 9).getCell (2).getStringCellValue ();
            poundalForSecondTable.add (poundalSecondTable);
        }

        ArrayList<String> unitOfMeasures = new ArrayList<> ();
        for (int i = 0; i < 5; i++) {
            String unitOfMeasure = sheet.getRow (i + 2).getCell (2).getStringCellValue ();
            unitOfMeasures.add (unitOfMeasure);
        }

        String expectedUnitOfAngle = "rad";
        String expectedPoundal = "N";
        String expectedUnitOfMeasure = "m";
        Assertions.assertAll (
                () -> Assertions.assertEquals (expectedPoundal, poundalForSecondTable.get (0)),
                () -> Assertions.assertEquals (expectedPoundal, poundalForSecondTable.get (1)),
                () -> Assertions.assertEquals (expectedPoundal, poundalForSecondTable.get (2)),
                () -> Assertions.assertEquals (expectedUnitOfMeasure, unitOfMeasures.get (0)),
                () -> Assertions.assertEquals (expectedUnitOfMeasure, unitOfMeasures.get (1)),
                () -> Assertions.assertEquals (expectedUnitOfMeasure, unitOfMeasures.get (2)),
                () -> Assertions.assertEquals (expectedPoundal, poundalForFirstTable),
                () -> Assertions.assertEquals (expectedUnitOfAngle, unitOfAngle)
        );
    }

    @Test
    public void checkStaticData() {
        StringBuilder staticDataFromFirstTable = new StringBuilder ();
        for (int i = 0; i < 6; i++) {
            String dataFromFirstTable = sheet.getRow (i + 1).getCell (0).getStringCellValue ();
            staticDataFromFirstTable.append (dataFromFirstTable).append (" ");
        }

        StringBuilder staticDataFromSecondTable = new StringBuilder ();
        for (int i = 0; i < 4; i++) {
            String dataFromSecondTable = sheet.getRow (i + 9).getCell (0).getStringCellValue ();
            staticDataFromSecondTable.append (dataFromSecondTable).append (" ");
        }

        String expectedStaticDataFromFirstTable = String.join (" ", "P1", "a", "b", "D1", "D2", "Beta");
        String expectedStaticDataFromSecondTable = String.join (" ", "P1y", "P1z", "P2y", "P2z");

        Assertions.assertAll (
                () -> Assertions.assertEquals (expectedStaticDataFromFirstTable, staticDataFromFirstTable.substring (0, staticDataFromFirstTable.length () - 1)),
                () -> Assertions.assertEquals (expectedStaticDataFromSecondTable, staticDataFromSecondTable.substring (0, staticDataFromSecondTable.length () - 1))
        );
    }

    @Test
    public void checkNamesHeaders() {
        ArrayList<String> namesFirstHeader = new ArrayList<> ();
        for (int i = 0; i < 3; i++) {
            String nameFirstHeader = sheet.getRow (0).getCell (i).getStringCellValue ();
            namesFirstHeader.add (nameFirstHeader);
        }

        ArrayList<String> namesSecondHeader = new ArrayList<> ();
        for (int i = 0; i < 3; i++) {
            String nameSecondHeader = sheet.getRow (8).getCell (i).getStringCellValue ();
            namesSecondHeader.add (nameSecondHeader);
        }

        String expectedNameFirstHeader = "Dane wejściowe, Wartość, Jednostka";
        String expectedNameSecondHeader = "Oznaczenie, Wartość, Jednostka";

        Assertions.assertAll (
                () -> Assertions.assertEquals (expectedNameFirstHeader, namesFirstHeader.get (0) + ", " + namesFirstHeader.get (1) + ", " + namesFirstHeader.get (2)),
                () -> Assertions.assertEquals (expectedNameSecondHeader, namesSecondHeader.get (0) + ", " + namesSecondHeader.get (1) + ", " + namesSecondHeader.get (2))
        );
    }

    @Test
    public void checkStyle() {
        HSSFCellStyle headerStyle = sheet.getRow (0).getCell (0).getCellStyle ();
        int foregroundColorHeader = headerStyle.getFillForegroundColor ();
        int alignmentHeader = headerStyle.getAlignment ();
        boolean isWrapTextHeader = headerStyle.getWrapText ();
        String fontNameHeader = headerStyle.getFont (workbook).getFontName ();
        int colorFontIndexHeader = headerStyle.getFont (workbook).getColor ();
        int boldWightFontHeader = headerStyle.getFont (workbook).getBoldweight ();
        boolean isItalicFontHeader = headerStyle.getFont (workbook).getItalic ();

        HSSFCellStyle dataStyle = sheet.getRow (1).getCell (0).getCellStyle ();
        String fontNameData = dataStyle.getFont (workbook).getFontName ();
        int boldWeightFontData = dataStyle.getFont (workbook).getBoldweight ();
        boolean isItalicFontData = dataStyle.getFont (workbook).getItalic ();

        int expectedForegroundColorHeader = 18;
        int expectedAlignmentHeader = 2;
        boolean expectedIsWrapTextHeader = true;
        String expectedFontNameHeader = "Arial";
        int expectedColorFontIndexHeader = 10;
        int expectedBoldWeightFontHeader = 700;
        boolean expectedIsItalicFontHeader = true;
        String expectedFontNameData = "Arial";
        int expectedBoldWeightFontData = 400;
        boolean expectedIsItalicFontData = false;

        Assertions.assertAll (
                () -> Assertions.assertTrue (foregroundColorHeader == expectedForegroundColorHeader),
                () -> Assertions.assertTrue (alignmentHeader == expectedAlignmentHeader),
                () -> Assertions.assertTrue (isWrapTextHeader == expectedIsWrapTextHeader),
                () -> Assertions.assertTrue (fontNameHeader.equals (expectedFontNameHeader)),
                () -> Assertions.assertTrue (colorFontIndexHeader == expectedColorFontIndexHeader),
                () -> Assertions.assertTrue (boldWightFontHeader == expectedBoldWeightFontHeader),
                () -> Assertions.assertTrue (isItalicFontHeader == expectedIsItalicFontHeader),
                () -> Assertions.assertTrue (fontNameData.equals (expectedFontNameData)),
                () -> Assertions.assertTrue (boldWeightFontData == expectedBoldWeightFontData),
                () -> Assertions.assertTrue (isItalicFontData == expectedIsItalicFontData)
        );
    }
}
