package com.java.excel.poi.dao;

import com.java.excel.poi.model.Book;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Book Dao implementation class
 */
public class BookDaoImpl implements IBookDao {

	private static final Logger LOGGER = Logger.getLogger(BookDaoImpl.class.getName());

	private static final String INSERT_QUERY = "INSERT INTO book(`title`, `author`, `price`) VALUES(?,?,?)";

	private static final String LIST_QUERY = "SELECT * FROM book";

	/**
	 * {@inheritDoc}
	 */
	@Override
	public boolean addBook(Book b) {

		boolean response = false;

		try(Connection connection = Manager.getConnection();
			PreparedStatement preparedStatement = connection.prepareStatement(INSERT_QUERY)){

			preparedStatement.setString(1, b.getTitle());
			preparedStatement.setString(2, b.getAuthor());
			preparedStatement.setDouble(3, b.getPrice());

			// execute insert SQL statement
			preparedStatement.executeUpdate();
			response = true;

		}catch(SQLException e){
			LOGGER.log(Level.SEVERE, String.format("Insert failure %s", e.getMessage()));

		}

		return response;
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public  List<Book> bookList(){


		List<Book> books = new ArrayList<>();
		try(Connection connection = Manager.getConnection();
			PreparedStatement preparedStatement = connection.prepareStatement(LIST_QUERY);
			ResultSet rs = preparedStatement.executeQuery()){

			while(rs.next()){
				Book b = new Book();
				b.setId((rs.getObject("id")!=null)?Integer.parseInt(rs.getObject("id").toString()):0);
				b.setTitle((rs.getObject("title")!=null)?rs.getObject("title").toString():null);
				b.setAuthor((rs.getObject("author")!=null)?rs.getObject("author").toString():null);
				b.setPrice((rs.getObject("price")!=null)?Double.parseDouble(rs.getObject("price").toString()):0);

				books.add(b);
			}

		}catch(SQLException e){
			LOGGER.log(Level.SEVERE, String.format("Get list failure %s", e.getMessage()));
		}

		return books;
	}

}
