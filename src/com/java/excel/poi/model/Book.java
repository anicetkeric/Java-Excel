package com.java.excel.poi.model;

import java.util.Objects;

/**
 * Book model
 */
public class Book {

	private int id;
	private String title;
	private String author;
	private double price;

	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public String getAuthor() {
		return author;
	}
	public void setAuthor(String author) {
		this.author = author;
	}
	public double getPrice() {
		return price;
	}
	public void setPrice(double price) {
		this.price = price;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		Book book = (Book) o;
		return id == book.id &&
				Double.compare(book.price, price) == 0 &&
				Objects.equals(title, book.title) &&
				Objects.equals(author, book.author);
	}

	@Override
	public int hashCode() {
		return Objects.hash(id, title, author, price);
	}

	@Override
	public String toString() {
		return "Book{" +
				"id=" + id +
				", title='" + title + '\'' +
				", author='" + author + '\'' +
				", price=" + price +
				'}';
	}
}
