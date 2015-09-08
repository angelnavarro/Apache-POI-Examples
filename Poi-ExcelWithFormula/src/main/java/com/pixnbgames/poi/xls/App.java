/**
 * Example of generation of an excel file that containts formulas with Apache POI
 */
package com.pixnbgames.poi.xls;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.pixnbgames.poi.xls.generator.ExcelGenerator;

/**
 * 
 * @author pixnbgames
 */
public class App {

	public static void main(String[] args) {
		HSSFWorkbook workbook = new ExcelGenerator().generateExcel();
		
		// Writing the excel to output file
		try {
			OutputStream out = new FileOutputStream("src/main/resources/ExcelWithFormula.xls");
			workbook.write(out);
			workbook.close();
			out.flush();
			out.close();
		}
		catch (IOException e) {
			System.err.println("Error at file writing");
			e.printStackTrace();
		}
	}

}
