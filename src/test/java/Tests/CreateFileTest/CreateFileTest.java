package Tests.CreateFileTest;

import Objects.CreateFile.CreateFileExcelCalculator;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;

public class CreateFileTest extends CreateFileBaseTest {

    @Test
    public void checkCreatedFile() {
        CreateFileExcelCalculator createFile = new CreateFileExcelCalculator ();
        createFile.createFileExcel (testDataReader.getTestData ().data (), configurationsReader.getPathFile ());

        File file = new File (configurationsReader.getPathFile ());

        Assertions.assertTrue (file.exists ());
    }
}
