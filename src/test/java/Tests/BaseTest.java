package Tests;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.jupiter.api.TestInstance;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public class BaseTest {
    HSSFWorkbook workbook;

    public void openFile() {
        try {
            FileInputStream file = new FileInputStream (new File ("C:\\Projects\\ProjectExcelFile\\test.xls"));
            this.workbook = new HSSFWorkbook (file);
            file.close ();
        } catch (FileNotFoundException e) {
            System.out.println (e.getMessage ());
        } catch (IOException io) {
            System.out.println ("Błąd odczytu");
        }
    }
}

