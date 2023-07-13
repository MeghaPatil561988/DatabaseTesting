

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class DriverFactory 
{
	public static WebDriver getWebDriver()
	{
		System.setProperty("webdriver.chrome.driver","C:\\\\Users\\\\Megha\\\\eclipse-workspace\\\\SeleniumJavaAutomation\\\\lib\\\\chromedriver.exe");
	    WebDriver driver = new ChromeDriver();
	    driver.manage().window().maximize();
	    return driver;


	}
	
}

