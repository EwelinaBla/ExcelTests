package Tests.CreateFileTest;

import Objects.CreateFile.CreateFile;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;

public class CreateFileTest extends CreateFileBaseTest {

    @Test
    public void checkCreatedFile()  {
        CreateFile createFile = new CreateFile ();
        createFile.createFileExcel (testDataReader.getTestData ().data ());

        File file=new File("C:\\Projects\\ProjectExcelFile\\calculator.xls");

        Assertions.assertTrue (file.exists ());
    }
}
