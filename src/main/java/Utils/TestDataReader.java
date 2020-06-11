package Utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class TestDataReader {
    Properties properties;
    private TestData testData;

    public void loadFile(String locationFile) {
        properties = new Properties();
        try {
            properties.load(new FileInputStream(locationFile));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void loadData() {
         testData = new TestData(properties);
    }

    public TestData getTestData() {
        return testData;
    }

}
