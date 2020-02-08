package com.google.tests;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHandler 
{
	File file;
	FileInputStream fis;
	FileOutputStream fout;
	Workbook wb;
	Sheet sheet;
	
	public ExcelHandler(String excelPath, String sheetName)
	{
		try 
		{
			file= new File(excelPath);
			fis= new FileInputStream(file);
			
			//Find the file extension by splitting file name in substring  and getting only extension name
		    String fileExtensionName = excelPath.substring(excelPath.indexOf("."));
		    
		    if(fileExtensionName.equals(".xlsx"))
		    {
		       //If it is xlsx file then create object of XSSFWorkbook class
		    	wb = new XSSFWorkbook(fis);
		    }
		    else if(fileExtensionName.equals(".xls"))
		    {
		       //If it is xls file then create object of XSSFWorkbook class
		    	wb = new HSSFWorkbook(fis);
		    }
		    
		    sheet=wb.getSheet(sheetName);
		} 
		
		catch (Exception e) 
		{
			System.out.println(e.getMessage());
		} 
	}
	
	public int getRowCount()
	{
		int rowCount= sheet.getLastRowNum();
		return(rowCount+1);
	}
	
	public int getColumnCount(int rowIndex)
	{
		int columnCount= sheet.getRow(rowIndex).getLastCellNum();
		return(columnCount);
	}
	
    public String getCellValue(int rowIndex,int colIndex)
    {
    	String cellValue=sheet.getRow(rowIndex).getCell(colIndex).getStringCellValue();
    	return cellValue;
    }

    public void addNewRow(int rowNum)
    //(String[] value,int colCount)
    {
    	//int rowCount=getRowCount();
    	//Row newrow=sheet.createRow(rowCount);
    	/*//int colCount=getColumnCount(0);
    	
    	for(int i=0; i<colCount;i++)
    	{
    		newrow.createCell(i).setCellValue(value[i]);
    	}*/
    	sheet=wb.getSheetAt(0);
    	sheet.createRow(rowNum);
    }
    
    public void updateCellValue(int rowIndex,int colIndex,String value)
    {
    	sheet.getRow(rowIndex).createCell(colIndex).setCellValue(value);
    }
    
    public void commit()
    {
    	try 
    	{
		    fout= new FileOutputStream(file);
			wb.write(fout);
			wb.close();
		} 
    	catch (Exception e) 
    	{
    		System.out.println(e.getMessage());
		}
    }
}
