import java.sql.Connection;

import java.sql.DriverManager;

import java.sql.SQLException;

import java.sql.Connection;

import java.sql.DriverManager;

import java.sql.ResultSet;

import java.sql.SQLException;

import java.sql.Statement;

import org.openqa.selenium.By;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Jdbcconnection 
{
	public static void main(String[] args) throws SQLException, ClassNotFoundException {

		// TODO Auto-generated method stub

		String host="localhost";

		String port= "3306";

		Connection con=DriverManager.getConnection("jdbc:mysql://" + host + ":" + port + "/selenium_practice", "root", "Swapnil88@");

		Statement s=con.createStatement();

		ResultSet rs=s.executeQuery("select * from login_table where scenario ='rewardscard'");

		while(rs.next())

		{

		WebDriver driver= new ChromeDriver();

		driver.get("https://login.salesforce.com");

		driver.findElement(By.xpath(".//*[@id='username']")).sendKeys(rs.getString("username"));

		driver.findElement(By.xpath(".//*[@id='password']")).sendKeys(rs.getString("password"));

		}

		}

}
