package org.excel;

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

public class NumericDateDemo {
	public static void main(String[]args) throws IOException {
        File f=new File ("C:\\Users\\Admin\\eclipse-workspace\\Project\\src\\test\\resources\\excel\\Project1.xlsx");
        FileInputStream f1=new FileInputStream(f);
    	Workbook w = new XSSFWorkbook(f1);
    	Sheet sheet= w.getSheet("Sheet1");
    	for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
    	Row row=sheet.getRow(i);
    	for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
    	Cell cell= row.getCell(j);
    	CellType cellType = cell.getCellType();
    	switch(cellType){
    		case STRING:
    			String stringCellValue = cell.getStringCellValue();
    			System.out.println(stringCellValue);
    			break;
    			default:
    				if(DateUtil.isCellDateFormatted(cell)) {
    					Date DateCellValue = cell.getDateCellValue();
    					///To define formate of data
    					SimpleDateFormat sd = new SimpleDateFormat("dd/mm/yyyy");
    					String Format = sd.format(DateCellValue);
    					System.out.println(Format);
    					break;
    					
    					}
    				else
    				{
    					double numericCellValue = cell.getNumericCellValue();
    					//to convert double to long
//    					long l = (long)numericCellValue;
//    					System.out.println(l);
    					BigDecimal valueof = BigDecimal.valueOf(numericCellValue);
    					String String = valueof.toString();
    					System.out.println(String);
    					break;
    					
    					
    				}
    			
    			
    	}
    

}
}
	}
}