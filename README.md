# Java-Excel - Maven Simple Project
Apache POI – Reading and Writing Excel file in Java
using Mysql

# Overview

Apache POI provides excellent support for working with Microsoft Excel documents. Apache POI is able to handle both XLS and XLSX formats of spreadsheets.

# MYSQL
 use book_db database
 
### Create database and tables
```sql
CREATE TABLE IF NOT EXISTS `book` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `title` varchar(35) NOT NULL,
  `author` varchar(35) NOT NULL,
  `price` double NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 AUTO_INCREMENT=1 ;
```

## JAVA
###Apache POI Maven Dependencies
```xml
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>3.17</version>
</dependency>

<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>3.17</version>
</dependency>

<!-- https://mvnrepository.com/artifact/mysql/mysql-connector-java -->
<dependency>
    <groupId>mysql</groupId>
    <artifactId>mysql-connector-java</artifactId>
    <version>8.0.22</version>
</dependency>


```
### Apache POI example program to read excel file to the list of book and insert in Mysql database.

```java
	
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
	
```

##### Output.java
```java

	    String excelFilePath = "book.xls";
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
	    
	   
```

### Apache POI example program to select records in MySql database to Write Excel File
```java

	
	
	public void writeBookListToFile(String fileName, List<Book> BookList) throws Exception{
		Workbook workbook = null;
		
		workbook=getWorkbook(fileName);
		
		
		if(workbook!=null){
			Sheet sheet = workbook.createSheet("Book");
			createHeaderRow(sheet);
			Iterator<Book> iterator = BookList.iterator();
			
			int rowIndex = 1;
			while(iterator.hasNext()){
				Book Book = iterator.next();
				Row row = sheet.createRow(rowIndex++);
				Cell cell0 = row.createCell(0);
				cell0.setCellValue(Book.getId());
				Cell cell1 = row.createCell(1);
				cell1.setCellValue(Book.getTitle());
				Cell cell2 = row.createCell(2);
				cell2.setCellValue(Book.getAuthor());
				Cell cell3 = row.createCell(3);
				cell3.setCellValue(Book.getPrice());
			}
			
			//lets write the excel data to file now
		    try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
		        workbook.write(outputStream);
		        outputStream.close();
		        System.out.println(fileName + " written successfully");	
		    }
				
		}
		
	}
	
	
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
	
	

```
##### Output.java
```java

metierBook=new BookDao();
	
		 try {
			 	
		List<Book> listp=metierBook.BookList();
		 Utiles reader = new Utiles();
			reader.writeBookListToFile("book.xlsx", listp);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

```

[Apache POI Example – Read and write Excel Formula](http://www.codejava.net/coding/working-with-formula-cells-in-excel-using-apache-poi)
