package com.pdf.convert;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConverterUtil {
	/**
	 * 设置excel宽度，保证转pdf后，可以显示在一页上
	 */
	public static void setExcelFitToWidth(String destFilePath) throws Exception{
		File file = new File(destFilePath);
		FileInputStream fis = new FileInputStream(file.getAbsolutePath());
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
		FileOutputStream fos = new FileOutputStream(file.getAbsolutePath());
		try{
			// 遍历sheet  设置sheet宽度自适应，确保可以打印在一页上
			for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
				XSSFSheet sheet = xssfWorkbook.getSheetAt(i);
				sheet.setPrintGridlines(true);
				sheet.setFitToPage(true);
				PrintSetup ps = sheet.getPrintSetup();
				ps.setFitWidth((short) 1);
				ps.setFitHeight((short) 0);
			}
			//写回文件
			xssfWorkbook.write(fos);	
		}finally{
			fos.close();
			xssfWorkbook.close();
			fis.close();
		}
	}
}
