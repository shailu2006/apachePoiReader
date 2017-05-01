package com.shailu.poi.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel2007Reader {

	public static void main(String[] args) {
		File file = new File("src/main/resources/Excel2007Employee.xlsx");
		Workbook workbook = null;

		try {
			workbook = WorkbookFactory.create(new FileInputStream(file));
			Sheet sheet = workbook.getSheetAt(0);
			Header header = sheet.getHeader();
			Row row;
			Cell cell;

			int rows = sheet.getLastRowNum();

			for (int i = 0; i < rows; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					short lastCellNum = row.getLastCellNum();
					for (int j = 0; j < lastCellNum; j++) {
						cell = row.getCell(j);
						if (cell != null) {
							System.out.println(cell.toString());
							System.out.println(cell.getCellStyle().getDataFormatString());
						}
					}
				}
				System.out.println("");
			}

		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} finally {
			try {
				// Close the workbook.
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
