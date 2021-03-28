package com.java.excel.poi.dao;


import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Manager  {

    private static final Logger LOGGER = Logger.getLogger(Manager.class.getName());

    public static final String URL = "jdbc:mysql://localhost:3306/book_db?autoReconnect=true&useSSL=false&serverTimezone=UTC";
    public static final String USER = "root";
    public static final String PASSWORD = "<YOUR_MYSQL_USER_PASSWORD>";
    public static final String DRIVER_CLASS = "com.mysql.cj.jdbc.Driver";
    private static final Manager INSTANCE = new Manager();

    private Manager() {
        try {
            Class.forName(DRIVER_CLASS);
        } catch (ClassNotFoundException e) {
            LOGGER.log(Level.SEVERE, e.getMessage());
        }
    }

    /**
     * Establish Java MySQL connection
     * @return Connection
     */
    private Connection createConnection() {

        Connection connection = null;
        try {
            connection = DriverManager.getConnection(URL, USER, PASSWORD);
        } catch (SQLException e) {
            LOGGER.log(Level.SEVERE, String.format("ERROR: Unable to Connect to Database: %s", e.getMessage()));
        }
        return connection;
    }   
     
    public static Connection getConnection() {
        return INSTANCE.createConnection();
    }
 	
}

