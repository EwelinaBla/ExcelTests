package Tests.CreateFileTest;

import Utils.TestDataReader;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.TestInstance;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public class CreateFileBaseTest {
    protected TestDataReader testDataReader;
    private final String locationFile = "src/test/TestData.properties";

    @BeforeAll
    public void configuration() {
        testDataReader = new TestDataReader();
        testDataReader.loadFile(locationFile);
        testDataReader.loadData();
    }

}
