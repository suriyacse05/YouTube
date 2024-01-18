package org.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Project1 {
	public static void main(String[]args) throws IOException {
          File f=new File ("C:\\Users\\Admin\\eclipse-workspace\\Project\\src\\test\\resources\\excel\\Project1.xlsx");
          FileInputStream f1=new FileInputStream(f);
      	Workbook w=new XSSFWorkbook(f1);
      	Sheet sheet= w.getSheet("Sheet1");
      	//to get all data
      	for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
      	Row row=sheet.getRow(i);
      	for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
      	Cell cell= row.getCell(j);
      	String stringcellvalue=cell.getStringCellValue();
      	if(stringcellvalue.equalsIgnoreCase("Test1")) {
      		cell.setCellValue("surya");
      		FileOutputStream f2 = new FileOutputStream(f);
      		w.write(f2);
      	}
      		
      	}
}
	}
}
