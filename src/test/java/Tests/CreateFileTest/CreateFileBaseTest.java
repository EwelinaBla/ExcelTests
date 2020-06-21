package Tests.CreateFileTest;

import Utils.ConfigurationsReader;
import Utils.TestDataReader;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.TestInstance;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public class CreateFileBaseTest {
    protected ConfigurationsReader configurationsReader;
    protected TestDataReader testDataReader;
    private final String configurationLocation = "src/Config/Configuration.properties";
    private final String locationFile = "src/test/TestData.properties";

    @BeforeAll
    public void configuration() {
        configurationsReader = new ConfigurationsReader (configurationLocation);
        testDataReader = new TestDataReader (locationFile);
    }
}
