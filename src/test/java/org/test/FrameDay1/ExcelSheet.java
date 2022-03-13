package org.test.FrameDay1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import io.github.bonigarcia.wdm.WebDriverManager;
public class ExcelSheet {
	public static void main(String[] args) throws IOException {
		WebDriverManager.chromedriver().setup();
		File file = new File("C:\\Users\\srira\\eclipse-workspace\\FrameDay1\\Excel\\Book1.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(stream);
		Sheet sheet = wb.getSheet("Datas");

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				CellType type = cell.getCellType();
				
				switch (type) {
				case STRING:
					
					String stringcellvalue = cell.getStringCellValue();
					System.out.println(stringcellvalue);
					
					break;
					
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						Date date1 = cell.getDateCellValue();
						SimpleDateFormat dateFormat2 = new SimpleDateFormat("dd/MMM/yyyy");
						String s = dateFormat2.format(date1);
						System.out.println(s);
						
					} else {
						double d1 = cell.getNumericCellValue();
						BigDecimal big = BigDecimal.valueOf(d1);
						String s2 = big.toString();
						System.out.println(s2);
				
					}
					break;
				default:
					break;	
			
				}

			}
		}
	}
}

		

