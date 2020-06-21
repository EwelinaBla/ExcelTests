package Utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public abstract class FileReader {
    Properties properties;

    public FileReader(String fileLocation) {
        loadFile (fileLocation);
        loadData ();
    }

    void loadFile(String locationFile) {
        properties = new Properties ();
        try {
            properties.load (new FileInputStream (locationFile));
        } catch (IOException e) {
            e.printStackTrace ();
            System.out.println("The file is incorrect or missing");
        }
    }

    abstract void loadData();
}
