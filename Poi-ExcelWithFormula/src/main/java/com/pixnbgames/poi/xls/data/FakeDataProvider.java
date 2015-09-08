package com.pixnbgames.poi.xls.data;

import java.util.ArrayList;
import java.util.List;

/**
 * Data provider for excel
 * 
 * @author angel
 */
public final class FakeDataProvider {
	
	/** Return the columns name for the table */
	public static List<String> getTableHeaders() {
		List<String> tableHeader = new ArrayList<String>();
		tableHeader.add("Product Name");
		tableHeader.add("Price");
		tableHeader.add("Amount");
		tableHeader.add("Subtotal");
		
		return tableHeader;
	}

	
	/**
	 * Return values for the table
	 * 
	 * @param numberOfRows Number of rows we want to receive
	 * 
	 * @return Values
	 */
	public static List<List<Object>> getTableContent(int numberOfRows) {
		if (numberOfRows <= 0) {
			throw new IllegalArgumentException("The number of rows must be a positive integer");
		}
		
		List<List<Object>> tableContent = new ArrayList<List<Object>>();

		List<Object> row = null;
		for (int i = 0; i < numberOfRows; i++) {
			tableContent.add(row = new ArrayList<Object>());
			row.add("Product " + i);
			row.add( (int)((Math.random() * 100.f) * 100) / 100.f);
			row.add( (int) (Math.random() * 10 + 1));
		}
		
		return tableContent;
	}
	
}
