package com.java.excel.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;





/**#################################################################
 * @author ANICET ERIC KOUAME
 * @Date: 23 févr. 2017
 * @Description:
 *Utiles
 *#################################################################*/
public class Utiles {
	
	
	public List<Book> readBookFromExcelFile(String excelFilePath) throws IOException {
	    List<Book> listb = new ArrayList<>();
	    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	 
	    Workbook workbook = new XSSFWorkbook(inputStream);
	    Sheet firstSheet = workbook.getSheetAt(0);
	    Iterator<Row> iterator = firstSheet.iterator();
	   
	    //get header number
	    Row header = firstSheet.getRow(0);
        int n = header.getLastCellNum();
        String headeridNumber = header.getCell(0).getStringCellValue();
        String headerTitleNumber = header.getCell(1).getStringCellValue();
        String headerAuthorNumber = header.getCell(2).getStringCellValue();
        String headerPriceNumber = header.getCell(3).getStringCellValue();
	   
        
        if (headerTitleNumber.equals("Title") && headerAuthorNumber.equals("Author") && headerPriceNumber.equals("Price")) {
        	

    	    //get the total number of columns in the file
    	    int noOfColumns = firstSheet.getRow(0).getPhysicalNumberOfCells();
    	    
    	    while (iterator.hasNext()) {
    	        Row nextRow = iterator.next();
    	        
    	    	// Not creating country object for header  
    	        if(nextRow.getRowNum()==0)  
    	         continue;
    	         
    	        Iterator<Cell> cellIterator = nextRow.cellIterator();
    	        Book b = new Book();
    	 
    	        while (cellIterator.hasNext()) {
    	            Cell nextCell = cellIterator.next();
    	            int columnIndex = nextCell.getColumnIndex();
    	           
    	          switch (columnIndex+1) {    	            	          
    	          case 2:
    	               b.setTitle((String) getCellValue(nextCell));
    	                break;
    	            case 3:
    	                b.setAuthor((String) getCellValue(nextCell));
    	                break;  
    	             case 4:
    	            	b.setPrice((Double) getCellValue(nextCell));
    	            	break; 
    	                   	                
    	            }
    	 
    	        }
    	        listb.add(b);
    	    }
        	
        	
        }else{
        	listb=null;
        }
        
        
	 
	    workbook.close();
	    inputStream.close();
	 
	    return listb;
	}
	
	@SuppressWarnings("unused")
	public Workbook getWorkbook(String excelFilePath)
	        throws IOException {
	    Workbook workbook = null;
	 
	    if (excelFilePath.endsWith("xlsx")) {
	        workbook = new XSSFWorkbook();
	    } else if (excelFilePath.endsWith("xls")) {
	        workbook = new HSSFWorkbook();
	    } else {
	        throw new IllegalArgumentException("The specified file is not Excel file");
	    }
	 
	    return workbook;
	}
	
	@SuppressWarnings("deprecation")
	private Object getCellValue(Cell cell) {
	    switch (cell.getCellType()) {
	    case Cell.CELL_TYPE_STRING:
	        return cell.getStringCellValue();
	 
	    case Cell.CELL_TYPE_BOOLEAN:
	        return cell.getBooleanCellValue();
	 
	    case Cell.CELL_TYPE_NUMERIC:
	        return cell.getNumericCellValue();
	    }
	 
	    return null;
	}
	
	
	
}
