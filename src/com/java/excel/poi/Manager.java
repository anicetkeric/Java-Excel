package com.java.excel.poi;


import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;



/**#################################################################
 * @author ANICET ERIC KOUAME
 * @Date: 23 févr. 2017
 * @Description:
 * Manager
 *#################################################################*/
public class Manager  {

	//static reference to itself
    private static Manager instance = new Manager();
    public static final String URL = "jdbc:mysql://localhost:3306/book_db";
    public static final String USER = "root";
    public static final String PASSWORD = "";
    public static final String DRIVER_CLASS = "com.mysql.jdbc.Driver"; 
     
    //private constructor
    private Manager() {
        try {
            //Step 2: Load MySQL Java driver
            Class.forName(DRIVER_CLASS);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }
     
    private Connection createConnection() {
 
        Connection connection = null;
        try {
            //Step 3: Establish Java MySQL connection
            connection = DriverManager.getConnection(URL, USER, PASSWORD);
        } catch (SQLException e) {
            System.out.println("ERROR: Unable to Connect to Database.");
        }
        return connection;
    }   
     
    public static Connection getConnection() {
        return instance.createConnection();
    }
 	
}

