package Utils;

import java.util.Properties;

public class TestData {
    private final String inputData;

    public TestData(Properties properties) {
        inputData = properties.getProperty ("inputData");
    }

    public String[] data() {
        String[] data = inputData.split (" ");
        return data;
    }
}