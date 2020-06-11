package Tests;


import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;

public class FileExcelTest extends BaseTest {
    private HSSFSheet sheet;

    @Test
    public void checkNameOfSheet() {
        openFile ();
        this.sheet = workbook.getSheetAt (0);
        String nameSheet = sheet.getSheetName ();

        Assertions.assertEquals ("Static", nameSheet);
    }

    @Test
    public void checkNumberOfSheets() {
        openFile ();
        int numberSheets = workbook.getNumberOfSheets ();

        Assertions.assertEquals (1, numberSheets);
    }

    @Test
    public void checkInputData() {
        openFile ();
        this.sheet = workbook.getSheetAt (0);
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
        openFile ();
        this.sheet = workbook.getSheetAt (0);
        String valueP1y = sheet.getRow (9).getCell (1).getCellFormula ();
        String expectedValueP1y = "ROUND((B2*COS(B7)),2)";

        String valueP1z = sheet.getRow (10).getCell (1).getCellFormula ();
        String expectedValueP1z = "ROUND((B2*SIN(B7)),2)";

        String valueP2y = sheet.getRow (11).getCell (1).getCellFormula ();
        String expectedValueP2y = "ROUND((B13*TAN(B7)),2)";

        String valueP2z = sheet.getRow (12).getCell (1).getCellFormula ();
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
        openFile ();
        this.sheet = workbook.getSheet ("Static");

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
        openFile ();
        this.sheet = workbook.getSheet ("Static");
        ArrayList<String> staticDataFromFirstTable = new ArrayList<> ();
        for (int i = 0; i < 6; i++) {
            String dataFromFirstTable = sheet.getRow (i + 1).getCell (0).getStringCellValue ();
            staticDataFromFirstTable.add (dataFromFirstTable);
        }

        ArrayList<String> staticDataFromSecondTable = new ArrayList<> ();
        for (int i = 0; i < 6; i++) {
            String dataFromFirstTable = sheet.getRow (i + 1).getCell (0).getStringCellValue ();
            staticDataFromSecondTable.add (dataFromFirstTable);
        }

        //var strings = List.of("foo", "bar", "baz");


        //List<String> strings = List.of ("P1y, P1z, P2y, P2z");

//        ArrayList<String> expextedStaticDataFromFirstTable= new ArrayList<>(List.of("P1, a, b, D1, D2, Beta"));
//        ArrayList<String> expextedStaticDataFromSecondTable= new ArrayList<>(List.of("P1y", "P1z", "P2y", "P2z"));P2z
//        ArrayList<String> friends =  new ArrayList<String>(Arrays.asList("Peter", "Paul"));
//

        ArrayList<String> places = Lists.newArrayList("Buenos Aires", "Córdoba", "La Plata");
        System.out.println (places);
//        ArrayList<String> strings = new ArrayList<>(List.of("foo", "bar"));
////        Assertions.assertAll (
//                ()-> Assertions.assertTrue (staticDataFromFirstTable.equals (expextedStaticDataFromFirstTable)),
//                ()-> Assertions.assertTrue (staticDataFromSecondTable.equals (expextedStaticDataFromSecondTable))
//        );
    }

    @Test
    public void checkNamesHeaders() {
        openFile ();
        this.sheet = workbook.getSheet ("Static");

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
        openFile ();
        this.sheet = workbook.getSheet ("Static");

        HSSFCellStyle headerStyle = sheet.getRow (0).getCell (0).getCellStyle ();
        int foregroundColorHeader = headerStyle.getFillForegroundColor ();
        int aligmantHeader=headerStyle.getAlignment ();
        boolean isWraptextHeader=headerStyle.getWrapText ();
        String fontNameHeader =headerStyle.getFont (workbook).getFontName ();
        int colorFontIndeksHeader = headerStyle.getFont (workbook).getColor ();
        int boldweightFontHeader = headerStyle.getFont (workbook).getBoldweight ();
        boolean isItalicFontHeader = headerStyle.getFont (workbook).getItalic ();


        HSSFCellStyle dataStyle = sheet.getRow (1).getCell (0).getCellStyle ();
        String fontNameData = dataStyle.getFont (workbook).getFontName ();
        int colorFontIndeksData = dataStyle.getFont (workbook).getBoldweight ();
        boolean isItalicFontData = dataStyle.getFont (workbook).getItalic ();

        int expectedForegroundColorHeader = 18;
        int expextedAligmantHeader = 2;
        boolean expectedIsWraptextHeader = true;
        String expectedFontNameHeader = "Arial";
        int expectedColorFontIndeksHeader = 10;
        int expectedBoldweightFontHeader = 700;
        boolean expectedIsItalicFontHeader = true;
        String expectedFontNameData = "Arial";
        int expectedColorFontIndeksData = 10;
        boolean expectedIsItalicFontData = false;

        Assertions.assertAll (
                ()-> Assertions.assertTrue (foregroundColorHeader==expectedForegroundColorHeader)
        );
    }
}


