package com.java.excel.poi.utils;

import com.java.excel.poi.model.Book;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.function.Function;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Excel file utils class
 */
public class ExcelFileUtils {

	private static final Logger LOGGER = Logger.getLogger(ExcelFileUtils.class.getName());

	private static final EnumMap<CellType, Function<Cell,Object>>
			EXCEL_CELL_VALUE = new EnumMap<>(CellType.class);

	static {
		EXCEL_CELL_VALUE.put(CellType.BLANK, null);
		EXCEL_CELL_VALUE.put(CellType.BOOLEAN, Cell::getBooleanCellValue);
		EXCEL_CELL_VALUE.put(CellType.STRING, Cell::getStringCellValue);
		EXCEL_CELL_VALUE.put(CellType.NUMERIC, cell -> {
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}
		});
	}

	/**
	 * read all row in Excel file
	 * @param excelFilePath path of file
	 * @return list of {@link Book}
	 */
	public List<Book> readBookFromExcelFile(String excelFilePath) {
		List<Book> bookList = new ArrayList<>();

		try (Workbook workbook = getWorkbook(excelFilePath)) {

			Sheet sheet = workbook.getSheetAt(0);

			//get header number
			Row header = sheet.getRow(0);
			String headeridNumber = header.getCell(0).getStringCellValue();
			String headerTitleNumber = header.getCell(1).getStringCellValue();
			String headerAuthorNumber = header.getCell(2).getStringCellValue();
			String headerPriceNumber = header.getCell(3).getStringCellValue();


			if (headeridNumber.equalsIgnoreCase("Id")
					&& headerTitleNumber.equalsIgnoreCase("Title")
					&& headerAuthorNumber.equalsIgnoreCase("Author")
					&& headerPriceNumber.equalsIgnoreCase("Price")) {

				sheet.forEach(row -> {
					// skip header
					if (row.getRowNum() != 0){
						bookList.add(transformer(row));
					}
				});

			} else {
				throw new ExcelFileException("File not compatible. please verify the columns name");
			}

			return bookList;
		} catch (IOException e) {
			throw new ExcelFileException("Cannot read the file: "+ e.getMessage());
		}
	}


	/**
	 * @param excelFilePath Excel file path
	 * @param books list of books
	 */
	public void writeBookListToFile(String excelFilePath, List<Book> books) {

		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet("Book");
			createHeaderRow(sheet);
			Iterator<Book> iterator = books.iterator();

			int rowIndex = 1;
			while(iterator.hasNext()){
				Book book = iterator.next();
				Row row = sheet.createRow(rowIndex++);
				Cell cell0 = row.createCell(0);
				cell0.setCellValue(book.getId());
				Cell cell1 = row.createCell(1);
				cell1.setCellValue(book.getTitle());
				Cell cell2 = row.createCell(2);
				cell2.setCellValue(book.getAuthor());
				Cell cell3 = row.createCell(3);
				cell3.setCellValue(book.getPrice());
			}

			//lets write the excel data to file now
			try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
				workbook.write(outputStream);
				LOGGER.log(Level.INFO, "File written successfully");
			}
		}
		catch (IOException e) {
		throw new ExcelFileException("Cannot read the file: "+ e.getMessage());
		}
	}


	/**
	 * Create header in Excel file
	 * @param sheet Excel sheet
	 */
	private void createHeaderRow(Sheet sheet) {

		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		Font font = sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) 12);
		cellStyle.setFont(font);


		cellStyle.setAlignment(HorizontalAlignment.CENTER);

		Row row = sheet.createRow(0);
		Cell cellId = row.createCell(0);
		cellId.setCellStyle(cellStyle);
		cellId.setCellValue("Id");

		Cell cellTitle = row.createCell(1);
		cellTitle.setCellStyle(cellStyle);
		cellTitle.setCellValue("Title");

		Cell cellAuthor = row.createCell(2);
		cellAuthor.setCellStyle(cellStyle);
		cellAuthor.setCellValue("Author");

		Cell cellPrice = row.createCell(3);
		cellPrice.setCellStyle(cellStyle);
		cellPrice.setCellValue("Price");
	}


	/**
	 * @param excelFilePath path of file
	 * @return Workbook
	 */
	private Workbook getWorkbook(String excelFilePath){
		Workbook workbook;

		try{
			try (FileInputStream inputStream = new FileInputStream(new File(excelFilePath))) {
				if (excelFilePath.endsWith("xlsx")) {
					workbook = new XSSFWorkbook(inputStream);
				} else {
					throw new ExcelFileException("The specified file is not Excel file");
				}
			}
		}catch(SecurityException | IOException  e){
			throw new ExcelFileException("The specified file is not Excel file");
		}

		return workbook;
	}


	/**
	 * transform excel row in Book object
	 * @param excelRoow excel file row
	 * @return object of {@link Book}
	 */
	private Book transformer(Row excelRoow){
		Book book = new Book();
		book.setTitle((String) getCellValue(excelRoow.getCell(1)));
		book.setAuthor((String) getCellValue(excelRoow.getCell(2)));
		book.setPrice((Double) getCellValue(excelRoow.getCell(3)));

		return book;
	}

	/**
	 * @param cell Row cell
	 * @return Object row
	 */
	private Object getCellValue(Cell cell) {
		try{
			return EXCEL_CELL_VALUE.get(cell.getCellTypeEnum()).apply(cell);
		}catch (Exception e){
			LOGGER.log(Level.WARNING, String.format("Cannot get cell value %s", e.getMessage()));
			return null;
		}
	}


}
