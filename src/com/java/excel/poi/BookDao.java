/**
 * 
 */
package com.java.excel.poi;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;





/**##########################################################
 * @author ANICET ERIC KOUAME
 * @Date 17 févr. 2017 08:55:18
 * @Description: 
 * @BookDao
 *#################################################################*/
public class BookDao implements BookI {

	
	@Override
	public int addBook(Book b) {
		
		String query = "INSERT INTO book"
				+ "(`title`, `author`, `price`) VALUES"
				+ "(?,?,?)";
		
		int out=0;
        Connection con = null;
        PreparedStatement preparedStatement = null;
        
		        try{
		        	con = Manager.getConnection();
		        	
		            
		            
		            preparedStatement = con.prepareStatement(query);

					preparedStatement.setString(1, b.getTitle());
					preparedStatement.setString(2, b.getAuthor());
					preparedStatement.setDouble(3, b.getPrice());

					// execute insert SQL stetement
					preparedStatement.executeUpdate();
		            out=1;
		            
		            
		        }catch(SQLException e){
			            e.printStackTrace();  
			            
			            out=0;
			        }finally{
			            try {
			                con.close();
			              
			            } catch (SQLException e) {
			                e.printStackTrace();
			                out=0;
			            }
			        }     
		       
		      
	return  out  ;
	}
	
		
}
