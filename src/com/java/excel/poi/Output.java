/**
 * 
 */
package com.java.excel.poi;

import java.io.IOException;
import java.util.List;



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
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		metierBook=new BookDao();
	
		
	    String excelFilePath = "book.xls";
	    Utiles reader = new Utiles();
	    List<Book> listp=null;
		
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
	    
		
					
			try {
				 	
					List<Book> list=metierBook.BookList();
					 Utiles writer = new Utiles();
					 writer.writeBookListToFile("book.xlsx", list);
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
			}
	

}
