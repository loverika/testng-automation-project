package com.zoho;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import io.github.bonigarcia.wdm.WebDriverManager;

public class zohoTest {
	
	WebDriver driver=null;
	WebDriverWait wait=null;
	String path="/Users/orcuncanlilar/selenium/testing-maven/selenium-maven-excel-zoho/src/test/resources/config.properties";
	Properties prop;
	String excelpath = "/users/orcuncanlilar/selenium/testing-maven/zohotestdata.xlsx";
	Xls_Reader data = new Xls_Reader(excelpath);
	Logger log = LogManager.getLogger(zohoTest.class.getName());
	
	
	@BeforeTest
	public void setUp() throws IOException {
		
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		FileInputStream ip=new FileInputStream(path);
		prop=new Properties();
		prop.load(ip);
		driver.get(prop.getProperty("url"));
		
	}
	
	@AfterTest
	public void closeUp() {
		driver.quit();
	}
	
	@Test
	public void printExcel() {
		
		Select s = new Select(driver.findElement(By.id("recPerPage")));
		s.selectByIndex(3);
		wait=new WebDriverWait(driver, 5);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("#reportTab>tbody>:nth-child(15)")));
		
		List<WebElement> rows = driver.findElements(By.cssSelector("#reportTab>tbody>tr>td:nth-child(1)"));
		List<WebElement> columns = driver.findElements(By.xpath("//table[@id='reportTab']/tbody/tr[1]/td"));
		
		List<WebElement> id = driver.findElements(By.cssSelector("#reportTab>tbody>tr>:nth-child(1)"));
		List<WebElement> name = driver.findElements(By.xpath("//table[@id='reportTab']/tbody/tr/td[2]"));
		List<WebElement> email = driver.findElements(By.xpath("//table[@id='reportTab']/tbody/tr/td[3]"));
		List<WebElement> phone = driver.findElements(By.xpath("//table[@id='reportTab']/tbody/tr/td[4]"));
		List<WebElement> salary = driver.findElements(By.xpath("//table[@id='reportTab']/tbody/tr/td[5]"));
		
		for(int i=0; i<rows.size(); i++) {
			data.setCellData("data", "EmployeeID", i+2, (id.get(i).getText()));
			log.debug("Writing EmployeID-"+(id.get(i).getText()));
			data.setCellData("data", "FullName", i+2, (name.get(i).getText()));
			log.debug("Writing FullName-"+(name.get(i).getText()));
			data.setCellData("data", "Email", i+2, (email.get(i).getText()));
			log.debug("Writing Email-"+(email.get(i).getText()));
			data.setCellData("data", "Phone", i+2, (phone.get(i).getText()));
			log.debug("Writing Phone-"+(phone.get(i).getText()));
			data.setCellData("data", "CurrentSalary", i+2, (salary.get(i).getText()));
			log.debug("Writing CurrentSalary-"+(salary.get(i).getText()));
		}
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
