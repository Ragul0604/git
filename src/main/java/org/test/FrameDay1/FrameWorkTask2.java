package org.test.FrameDay1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FrameWorkTask2 {
	public static void main(String[] args) throws IOException {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("http://demo.automationtesting.in/Register.html");
		driver.manage().window().maximize();
		
		WebElement skills = driver.findElement(By.id("Skills"));
		
		Select select = new Select(skills);
		
		List<WebElement> options = select.getOptions();
		
		File file = new File("C:\\Users\\srira\\eclipse-workspace\\FrameDay1\\Excel\\Frameworktask2.xlsx");
		Workbook b = new XSSFWorkbook();
		Sheet sheet = b.createSheet();
		
		int i=0;
		for (WebElement x : options) {
			String text = x.getText();
			Row row = sheet.createRow(i);
			Cell cell = row.createCell(0);
			cell.setCellValue(text);
			i++;
			
		}
		FileOutputStream stream = new FileOutputStream(file);
		b.write(stream);
		
		
		
		
		
		
		
	}

}
