package com.java.excel.poi;

import com.java.excel.poi.dao.BookDaoImpl;
import com.java.excel.poi.dao.IBookDao;
import com.java.excel.poi.model.Book;
import com.java.excel.poi.utils.ExcelFileUtils;

import java.util.List;


public class Output {

	private static final String EXCEL_FILE_INPUT_PATH = "book.xlsx";
	private static final String EXCEL_FILE_OUTPUT_PATH = "book-java.xlsx";
	private static final IBookDao bookDao = new BookDaoImpl();

	public static void main(String[] args) {

		readFile();

		writeFile();
	}

	/**
	 * read data in database and create new file
	 */
	private static void writeFile() {
		List<Book> books = bookDao.bookList();
		ExcelFileUtils writer = new ExcelFileUtils();
		writer.writeBookListToFile(EXCEL_FILE_OUTPUT_PATH, books);
	}

	/**
	 * read datas in Excel file
	 */
	private static void readFile() {

		ExcelFileUtils reader = new ExcelFileUtils();
		List<Book> books;

		// get All row in file
		books = reader.readBookFromExcelFile(EXCEL_FILE_INPUT_PATH);

		// inserting into the database
		books.forEach(bookDao::addBook);
	}
	

}
