package Tests;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;

import java.io.*;

public class BaseTest {
    HSSFWorkbook workbook;

    @BeforeEach
    public void openFile() {
        try {
            FileInputStream file = new FileInputStream (new File ("C:\\Projects\\ProjectExcelFile\\test.xls"));
            this.workbook = new HSSFWorkbook (file);
        } catch (FileNotFoundException e) {
            System.out.println (e.getMessage ());
        } catch (IOException io) {
            System.out.println ("Błąd odczytu");
            System.exit (1);
        }
    }
    @AfterEach
    public void closeFile() {
        try {
            FileOutputStream out = new FileOutputStream (new File ("C:\\Projects\\ProjectExcelFile\\test.xls"));
            workbook.write (out);
            out.close ();
        } catch (FileNotFoundException e) {
            System.out.println (e.getMessage ());
        } catch (IOException io) {
            System.out.println ("Błąd zapisu");
            System.exit (1);
        }
    }
}

