package com.charts;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RetrieveProjectName 
{
	public static void main(String args[])
	{
		RetrieveProjectName obj=new RetrieveProjectName();
		obj.getNames();
	}
	public ArrayList<String> getNames()
	{
		

		ArrayList projects=new ArrayList();
		try
		{
			InputStream fis = this.getClass().getClassLoader().getResourceAsStream("Team_GitSheet.xlsx");
				      XSSFWorkbook workbook = new XSSFWorkbook(fis);
				      XSSFSheet spreadsheet = workbook.getSheet("Sheet1");
				      int noofrows=spreadsheet.getLastRowNum();
				      System.out.println("no of rows is"+noofrows);
				      for (int i=1; i<=spreadsheet.getLastRowNum();i++)
				      {
				    	  XSSFRow row1 = spreadsheet.getRow(i);
				    	  XSSFCell cell1=row1.getCell(3);
				    	  String name=cell1.toString();
				    	  if (!projects.contains(name))
				    	  {
				    	  projects.add(name);  
				    	  }
				          
				    	  
				      }
				      System.out.println(projects);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		return projects;
	}
}
