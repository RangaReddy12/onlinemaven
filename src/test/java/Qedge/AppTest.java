package Qedge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class AppTest {
WebDriver driver;
FileInputStream fi;
FileOutputStream fo;
XSSFWorkbook wb;
XSSFSheet ws;
XSSFRow row;
File screen;
String inputpath="E:\\LoginData.xlsx";
String outputpath="E:\\Results.xlsx";
@BeforeTest
public void setup()
{
System.setProperty("webdriver.chrome.driver", "E:\\Online_Framework\\Orange_HRM\\CommonDrivers\\chromedriver.exe");
driver=new ChromeDriver();
}
@Test
public void verifyLogin()throws Throwable
{
//access path of excel
fi=new FileInputStream(inputpath);
//get wb from file
wb=new XSSFWorkbook(fi);
//get sheet from wb
ws=wb.getSheet("Login");
//get first row from sheet
row=ws.getRow(0);
//count no of rows in sheet
int rc=ws.getLastRowNum();
int cc=row.getLastCellNum();
Reporter.log("no of rows are::"+rc+"  "+"no of columns in first row::"+cc,true);
for(int i=1;i<=rc;i++)
{
driver.get("http://orangehrm.qedgetech.com/");	
driver.manage().window().maximize();
//get username column data
String username=ws.getRow(i).getCell(0).getStringCellValue();
String password=ws.getRow(i).getCell(1).getStringCellValue();
//fill login form
driver.findElement(By.name("txtUsername")).sendKeys(username);
driver.findElement(By.name("txtPassword")).sendKeys(password);
driver.findElement(By.name("Submit")).click();
if(driver.getCurrentUrl().contains("dash"))
{
	//take screen shot when test pass
screen=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
FileUtils.copyFile(screen, new File("E:\\Online_Framework\\Orange_HRM\\Screen\\"+i+"Loginpage.png"));
Reporter.log("Login Success",true);
//write login success into results column
ws.getRow(i).createCell(2).setCellValue("Login Success");
//write pass  into status column
ws.getRow(i).createCell(3).setCellValue("Pass");
}
else
{
	//take screen shot when test pass
screen=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
FileUtils.copyFile(screen, new File("E:\\Online_Framework\\Orange_HRM\\Screen\\"+i+"Loginpage.png"));	
//get error message and store
String message=driver.findElement(By.id("spanMessage")).getText();
Reporter.log(message,true);
//write message into results column
ws.getRow(i).createCell(2).setCellValue(message);
ws.getRow(i).createCell(3).setCellValue("Fail");
}
}
fi.close();
fo=new FileOutputStream(outputpath);
wb.write(fo);
fo.close();
wb.close();
}
@AfterTest
public void teardown()
{
	driver.close();
}
}









