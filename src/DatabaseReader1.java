

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class DatabaseReader1 
{
	public static void main(String[] args) throws ClassNotFoundException, SQLException 
	{
		 String driverClassName="com.mysql.cj.jdbc.Driver";
		 String url = "jdbc:mysql://localhost:3306/selenium_practice";   
		 String dbUserName = "root";
	        String dbPassword  = "Swapnil88@";
	 
	        Class.forName(driverClassName);
	        try {
	        Connection con = DriverManager.getConnection(url, dbUserName, dbPassword);
	        
	        String sql = "SELECT * FROM login_table;";
	        PreparedStatement pst = con.prepareStatement(sql);
	        ResultSet rs = pst.executeQuery();
	        while (rs.next()) {
	            String userName = rs.getString("user_name");
	            String password = rs.getString("password");
	            login(userName, password);
	        System.out.println("connected to the database"); 
	        }
	        }
	        catch(SQLException e)
	        {
	        	System.out.println("error");
	        	e.printStackTrace();
	        }
	 
	}
	
	
	
	 public static void login(String userName, String password)
     {
     	 // Step1: Open Web Browser using DriverFactory
         WebDriver driver1 = DriverFactory.getWebDriver();
  
         // Step2: Open the the web-application URL: https://www.saucedemo.com/
         String url = "https://www.saucedemo.com/";
         driver1.get(url);
  
         // Step3: call the login function using variables String userName and password
  
         // Step3: Enter userName
         WebElement userNameWebElement = driver1.findElement(By.id("user-name"));
         userNameWebElement.sendKeys(userName);
  
         // Step4: Enter password
         WebElement passwordWebElent = driver1.findElement(By.id("password"));
         passwordWebElent.sendKeys(password);
  
         // Step5: Click on LOGIN button
         WebElement loginButtonWebElement = driver1.findElement(By.id("login-button"));
         loginButtonWebElement.click();
         
        // System.out.println("login times"+userName);
        // System.out.println("login times"+password);
         driver1.quit();
  
     }

}
