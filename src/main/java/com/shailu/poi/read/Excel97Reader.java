package com.shailu.poi.read;

import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class Excel97Reader {

	public static void main(String[] args) {
		File file = new File("src/main/resources/Excel97Employee.xls");
		HSSFWorkbook workbook = null;

		try {
			POIFSFileSystem poifsFileSystem = new POIFSFileSystem(file);
			workbook = new HSSFWorkbook(poifsFileSystem);
			HSSFSheet sheet = workbook.getSheetAt(0);
			HSSFRow row;
			HSSFCell cell;

			int rows = sheet.getPhysicalNumberOfRows();

			for (int i = 0; i < rows; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					int physicalNumberOfCells = row.getPhysicalNumberOfCells();
					for (int j = 0; j < physicalNumberOfCells; j++) {
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
