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
		tableHeader.add("Name");
		tableHeader.add("Address");
		tableHeader.add("Phone");
		
		return tableHeader;
	}

	
	/**
	 * Return values for the table
	 * 
	 * @param numberOfRows Number of rows we want to receive
	 * 
	 * @return Values
	 */
	public static List<List<String>> getTableContent(int numberOfRows) {
		if (numberOfRows <= 0) {
			throw new IllegalArgumentException("The number of rows must be a positive integer");
		}
		
		List<List<String>> tableContent = new ArrayList<List<String>>();

		List<String> row = null;
		for (int i = 0; i < numberOfRows; i++) {
			tableContent.add(row = new ArrayList<String>());
			row.add("my name is " + i);
			row.add("my address is " + i);
			row.add("my phone is " + i);
		}
		
		return tableContent;
	}
	
}
