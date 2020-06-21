package Utils;

public class TestDataReader extends FileReader {
    private TestData testData;
    private String fileLocation;

    public TestDataReader(String fileLocation) {
        super (fileLocation);
        this.fileLocation = getFileLocation ();
    }

    void loadData() {
        testData = new TestData (properties);
    }

    public TestData getTestData() {
        return testData;
    }

    public String getFileLocation() {
        return fileLocation;
    }
}
