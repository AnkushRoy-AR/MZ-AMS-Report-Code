import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class PropertiesLoader {

	Properties prop = new Properties();
	
	public PropertiesLoader() throws FileNotFoundException, IOException
	{
		InputStream input;
		try {
			input = new FileInputStream("report.properties");
			prop.load(input);
		}
		catch (FileNotFoundException e)
		{
			e.printStackTrace();
			throw e;
		}
		catch(IOException e)
		{
			e.printStackTrace();
			throw e;
		}
	}
	
	public String getValue(String key)
	{
		String value = prop.getProperty(key);
		return value;
	}
}
