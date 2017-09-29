package tw.jimmy.lab;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ReadXLS {
	public static void main(String[] args) {
		HSSFWorkbook readWorkbook;
		FileInputStream fis;
		try {
			File file = new File("./files/test.xls");
			fis = new FileInputStream(file);
			readWorkbook = new HSSFWorkbook(fis);
			// 取得總Sheet數
			// int sNo = readWorkbook.getNumberOfSheets();
			// 指定sheet的名稱,取得Sheet
			// readWorkbook.getSheetIndex(name);
			// 指定Sheet的位置取得Sheet
			HSSFSheet readSheet = readWorkbook.getSheetAt(0);
			// 取得列總數
			// int rNo = readSheet.getPhysicalNumberOfRows();
			// 先取出列
			HSSFRow r = readSheet.getRow(0);
			// 再取出欄
			System.out.println(r.getCell(0).getNumericCellValue() + " "
					+ r.getCell(1).getNumericCellValue());
			readWorkbook.close();
			fis.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
