package com.java.excel.poi.dao;

import com.java.excel.poi.model.Book;

import java.util.List;

/**
 * Book Dao interface
 */
public interface IBookDao {

	/**
	 * Insert new book
	 * @param book instance of {@link Book}
	 * @return true if successfully otherwise false
	 */
	boolean addBook(Book book);

	/**
	 * Get list of book
	 * @return list of {@link Book}
	 */
	List<Book> bookList();
}
