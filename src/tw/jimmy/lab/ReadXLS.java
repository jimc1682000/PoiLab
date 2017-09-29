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
			// ���o�`Sheet��
			// int sNo = readWorkbook.getNumberOfSheets();
			// ���wsheet���W��,���oSheet
			// readWorkbook.getSheetIndex(name);
			// ���wSheet����m���oSheet
			HSSFSheet readSheet = readWorkbook.getSheetAt(0);
			// ���o�C�`��
			// int rNo = readSheet.getPhysicalNumberOfRows();
			// �����X�C
			HSSFRow r = readSheet.getRow(0);
			// �A���X��
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
