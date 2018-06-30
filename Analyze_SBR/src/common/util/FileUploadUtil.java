package common.util;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileUploadUtil {

	public String readExcelFile(String filePath) {
		File excelFile = new File(filePath);
		String resultVaue = null;

		if (excelFile.exists()) {
			if (filePath.contains(".xls") || filePath.contains(".XLS")) {
				xlsRead(filePath);
			} else if (filePath.contains(".xlsx") || filePath.contains(".XLSX")) {
				xlsxRead(filePath);
			} else {
				resultVaue = "false";
			}
		} else {
			resultVaue = "false";
		}
		return resultVaue;
	}

	public List<Map<String, String>> xlsRead(String filePath) {
		List<Map<String, String>> resultListMap = null;
		try {
			FileInputStream fis = new FileInputStream(filePath);
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheetAt(0);
			
			int rowindex = 0;			
			int rows = sheet.getPhysicalNumberOfRows();
			
			for (rowindex = 1; rowindex < rows; rowindex++) {
				HSSFRow row = sheet.getRow(rowindex);
				if (row != null) {
					int cells = row.getPhysicalNumberOfCells();
					if(cells != 0){
						Map<String, String> resultValue = new HashMap();
						HSSFCell cell_scan_no = row.getCell(1);
						HSSFCell cell_site_url= row.getCell(2);
						resultValue.put("scan_no", cell_scan_no.getStringCellValue());
						resultValue.put("board_no", Integer.toString(rowindex));
						resultValue.put("site_url", cell_site_url.getStringCellValue());
						resultListMap.add(resultValue);
					}else{
						continue;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return resultListMap;
	}

	public List<Map<String, String>> xlsxRead(String filePath) {
		List<Map<String, String>> resultListMap = null;
		try {
			FileInputStream fis=new FileInputStream(filePath);
			XSSFWorkbook workbook=new XSSFWorkbook(fis);
			XSSFSheet sheet=workbook.getSheetAt(0);
			
			int rowindex = 0;			
			int rows = sheet.getPhysicalNumberOfRows();
			
			for (rowindex = 1; rowindex < rows; rowindex++) {
			    XSSFRow row=sheet.getRow(rowindex);
				if (row != null) {
					int cells = row.getPhysicalNumberOfCells();
					if(cells != 0){
						Map<String, String> resultValue = new HashMap();
						XSSFCell cell_scan_no = row.getCell(1);
						XSSFCell cell_site_url= row.getCell(2);
						resultValue.put("scan_no", cell_scan_no.getStringCellValue());
						resultValue.put("board_no", Integer.toString(rowindex));
						resultValue.put("site_url", cell_site_url.getStringCellValue());
						resultListMap.add(resultValue);
					}else{
						continue;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return resultListMap;
	}
}
