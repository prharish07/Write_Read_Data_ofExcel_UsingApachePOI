package com.p4programmer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo1 {
	

	public void writeMyDataToExcel()
	{
		//Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); 
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Salary");
          
        //This data needs to be written (Object[])
       // This Data List also can be the list of rows from  table of DB
        
        
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "NAME", "Salary","Age"});
        data.put("2", new Object[] {1, "HARISH", "5000",26});
        data.put("3", new Object[] {2, "LokuDasu", "25000",25});
        data.put("4", new Object[] {3, "Pavani", "26000",24});
        data.put("5", new Object[] {4, "Anil", "30000",24});
        data.put("6", new Object[] {6, "Neha Gautham", "45000",26});
        data.put("7", new Object[] {7, "Manoj KP", "35000",27});
        
          
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);  // if the column cell contain String value
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);  // if the cell value is integer
            }
        }
        try
        {
            //Write the workbook in file system
        	//Location is our Wish where we want to place it.
        	
            FileOutputStream out = new FileOutputStream(new File("MyEmpSalaries_demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("MyEmpSalaries_demo.xlsx written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
	}
	
	@SuppressWarnings("deprecation")
	public void  readMyDataFromExcel()
	{
		try
        {
			//This file also can be given as a options to upload/browse from the user interface .
			//i had just used the above generated file as the read file 
			
            FileInputStream file = new FileInputStream(new File("MyEmpSalaries_demo.xlsx"));
 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) 
                    {
                        case Cell.CELL_TYPE_NUMERIC: //Integer Cell Type
                            System.out.print(cell.getNumericCellValue() + "\t\t"); 
                            break;
                        case Cell.CELL_TYPE_STRING: //String Cell Type
                            System.out.print(cell.getStringCellValue() + "\t\t");
                            break;
                    }
                }
                System.out.println("Its Done.... ");
            }
            file.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
	}
    public static void main(String[] args) 
    {
        ReadExcelDemo1 rdObj=new ReadExcelDemo1();
       // rdObj.writeMyDataToExcel();
        System.out.println(" Now We are Going to read Excel File .... \n ");
        
        rdObj.readMyDataFromExcel();
        
    }

}
