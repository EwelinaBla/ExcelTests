package Utils;

public class ConfigurationsReader extends FileReader {

    private String pathFile;
    private String configurationLocation;

    public ConfigurationsReader(String configurationLocation) {
        super (configurationLocation);
        this.configurationLocation = getConfigurationLocation ();
    }

    void loadData() {
        pathFile = properties.getProperty ("pathFile");
    }

    public String getConfigurationLocation() {
        return configurationLocation;
    }

    public String getPathFile() {
        return pathFile;
    }
}
