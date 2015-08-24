package com.pixnbgames.poi.xls.generator;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import com.pixnbgames.poi.xls.data.FakeDataProvider;

public class ExcelGenerator {
	
	// Excel work book
	private HSSFWorkbook workbook;
	
	// Fonts
	private HSSFFont headerFont;
	private HSSFFont contentFont;
	
	// Styles
	private HSSFCellStyle headerStyle;
	private HSSFCellStyle oddRowStyle;
	private HSSFCellStyle evenRowStyle;
	
	// Integer to store the index of the next row
	private int rowIndex;
	

	/**
	 * Make a new excel workbook with sheet that contains a stylized table
	 * 
	 * @return
	 */
	public HSSFWorkbook generateExcel() {
		
		// Initialize rowIndex
		rowIndex = 0;
		
		// New Workbook
		workbook = new HSSFWorkbook();
		
		// Generate fonts
		headerFont  = createFont(HSSFColor.WHITE.index, (short)12, true);
		contentFont = createFont(HSSFColor.BLACK.index, (short)10, false);
		
		// Generate styles
		headerStyle  = createStyle(headerFont,  HSSFCellStyle.ALIGN_CENTER, HSSFColor.BLUE_GREY.index,       true, HSSFColor.WHITE.index);
		oddRowStyle  = createStyle(contentFont, HSSFCellStyle.ALIGN_LEFT,   HSSFColor.GREY_25_PERCENT.index, true, HSSFColor.GREY_80_PERCENT.index);
		evenRowStyle = createStyle(contentFont, HSSFCellStyle.ALIGN_LEFT,   HSSFColor.GREY_40_PERCENT.index, true, HSSFColor.GREY_80_PERCENT.index);
		
		// New sheet
		HSSFSheet sheet = workbook.createSheet("Very Cool Sheet");
		
		// Table header
		HSSFRow      headerRow    = sheet.createRow( rowIndex++ );
		List<String> headerValues = FakeDataProvider.getTableHeaders();
		
		HSSFCell headerCell = null;
		for (int i = 0; i < headerValues.size(); i++) {
			headerCell = headerRow.createCell(i);
			headerCell.setCellStyle(headerStyle);
			headerCell.setCellValue( headerValues.get(i) );
		}
		
		
		// Table content
		HSSFRow  contentRow  = null;
		HSSFCell contentCell = null;
		
		// Obtain table content values
		List<List<String>> contentRowValues = FakeDataProvider.getTableContent(20);
		for (List<String> rowValues : contentRowValues) {
			
			// At each row creation, rowIndex must grow one unit
			contentRow = sheet.createRow( rowIndex++ );
			for (int i = 0; i < rowValues.size(); i++) {
				contentCell = contentRow.createCell(i);
				contentCell.setCellValue( rowValues.get(i) );
				
				// Style depends on if row is odd or even
				contentCell.setCellStyle( rowIndex % 2 == 0 ? oddRowStyle : evenRowStyle );
			}
		}
		
		
		// Autosize columns
		for (int i = 0; i < headerValues.size(); sheet.autoSizeColumn(i++));
		
		return workbook;
	}
	
	
	/**
	 * Create a new font on base workbook
	 * 
	 * @param fontColor       Font color (see {@link HSSFColor})
	 * @param fontHeight      Font height in points
	 * @param fontBold        Font is boldweight (<code>true</code>) or not (<code>false</code>)
	 * 
	 * @return New cell style
	 */
	private HSSFFont createFont(short fontColor, short fontHeight, boolean fontBold) {
		
		HSSFFont font = workbook.createFont();
		font.setBold(fontBold);
		font.setColor(fontColor);
		font.setFontName("Arial");
		font.setFontHeightInPoints(fontHeight);
		
		return font;
	}
	
	
	/**
	 * Create a style on base workbook
	 * 
	 * @param font            Font used by the style
	 * @param cellAlign       Cell alignment for contained text (see {@link HSSFCellStyle})
	 * @param cellColor       Cell background color (see {@link HSSFColor})
	 * @param cellBorder      Cell has border (<code>true</code>) or not (<code>false</code>)
	 * @param cellBorderColor Cell border color (see {@link HSSFColor})
	 * 
	 * @return New cell style
	 */
	private HSSFCellStyle createStyle(HSSFFont font, short cellAlign, short cellColor, boolean cellBorder, short cellBorderColor) {

		HSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);
		style.setAlignment(cellAlign);
		style.setFillForegroundColor(cellColor);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		
		if (cellBorder) {
			style.setBorderTop(HSSFCellStyle.BORDER_THIN);
			style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			style.setBorderRight(HSSFCellStyle.BORDER_THIN);
			style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			
			style.setTopBorderColor(cellBorderColor);
			style.setLeftBorderColor(cellBorderColor);
			style.setRightBorderColor(cellBorderColor);
			style.setBottomBorderColor(cellBorderColor);
		}
		
		return style;
	}
}
