package orangee;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class AddEmployee {
	public WebDriver dr;
	public XSSFWorkbook wb;
	public XSSFSheet sh;
	public FileOutputStream fo;
	public FileInputStream fi;
	public File f;
	@BeforeTest
	public void Launch() throws InterruptedException{
	//	System.setProperty("webdriver.gecko.driver", "‪D:\\Drivers\\geckodriver.exe");
		//dr=new ChromeDriver();
		dr=new FirefoxDriver();
		//dr.manage().window().maximize();
		dr.get("http://opensource.demo.orangehrmlive.com");
		dr.findElement(By.xpath("html/body/div[1]/div/div[2]/form/div[2]/input")).sendKeys("Admin");
		dr.findElement(By.name("txtPassword")).sendKeys("admin");
		dr.findElement(By.xpath("html/body/div[1]/div/div[2]/form/div[5]/input")).submit();
		Thread.sleep(3000);
		
	}
@Test
public void login() throws InterruptedException, IOException{
	fi=new FileInputStream("D:\\Selenium_Framework\\OrangeHRM\\Employees.xlsx");
	wb=new XSSFWorkbook(fi);
	sh=wb.getSheetAt(0);
	int rc=sh.getLastRowNum();
	System.out.println(rc);
	for (int i = 1; i < rc; i++) {
		
		
	String firstname=sh.getRow(i).getCell(0).getStringCellValue();
	String lastname=sh.getRow(i).getCell(1).getStringCellValue();
	String employeeid=sh.getRow(i).getCell(2).getStringCellValue();
	
	dr.findElement(By.id("menu_pim_viewPimModule")).click();
	Thread.sleep(3000);
	dr.findElement(By.id("menu_pim_addEmployee")).click();
	Thread.sleep(2000);
	dr.findElement(By.name("firstName")).sendKeys(firstname);
	dr.findElement(By.name("lastName")).sendKeys(lastname);
	dr.findElement(By.id("employeeId")).clear();
	Thread.sleep(2000);
	dr.findElement(By.id("employeeId")).sendKeys(employeeid);
	dr.findElement(By.id("btnSave")).click();
	Thread.sleep(2000);
	String Expval="OrangeHRM";
	String Acval=dr.getTitle();
	if(Expval.equals(Acval))
	{
		Reporter.log("complete",true);
		sh.getRow(i).createCell(3).setCellValue("complete");
	}
	else{
		Reporter.log("incomplete",true);
		sh.getRow(i).createCell(3).setCellValue("incomplete");
	}
	
	dr.findElement(By.id("menu_pim_addEmployee")).click();
	}
	Date d=new Date();
	DateFormat df=new SimpleDateFormat("dd_MM_yyyy HH_mm_ss");
	String date=df.format(d);
	File src=((TakesScreenshot)dr).getScreenshotAs(OutputType.FILE);
	FileUtils.copyFile(src, new File("D:\\Screenshot\\"+date+" "+"orangeEmployee.jpg"));
	
	
	
	
	
}
@AfterTest
public void logout(){
	dr.quit();
}
}
