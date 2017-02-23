/**
 * 
 */
package com.java.excel.poi;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;


/**#################################################################
 * @author ANICET ERIC KOUAME
 * @Date: 23 févr. 2017
 * @Description:
 *Output
 *#################################################################*/

public class Output {

	private static BookI metierBook;
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub

metierBook=new BookDao();
		
		/*Book b=new Book();
		b.setTitle("Amor");
		b.setAuthor("Anthony");
		b.setPrice(5000);*/
		
		//System.out.println(metierBook.addBook(b));
		
		
	    String excelFilePath = "book.xlsx";
	    Utiles reader = new Utiles();
	    List<Book> listp=null;
		
	    	try {
				
	    		
	    		if(reader.getWorkbook(excelFilePath)!=null){
					
		    			
		    			 try {
		    					
		    					listp = reader.readBookFromExcelFile(excelFilePath);
		    					if(listp.size()>0){
		    						
		    						for (Book book : listp) {
		    							metierBook.addBook(book);
		    						}
		    						
		    						System.out.println("Reading the excel file and inserting into the database");
		    						
		    					}else System.out.println("File not compatible");
		    					
		    				} catch (IOException e) {
		    					// TODO Auto-generated catch block
		    					e.printStackTrace();
		    					System.out.println("Error: "+ e.getMessage());
		    				}	
						}
			
	    	
	    	} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
	    
	   
	    
		
		
	}

}
