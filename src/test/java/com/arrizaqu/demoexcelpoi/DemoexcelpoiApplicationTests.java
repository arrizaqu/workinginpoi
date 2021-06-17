package com.arrizaqu.demoexcelpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class DemoexcelpoiApplicationTests {

	@Test
	void contextLoads() {
	}
	
	@Test
	void testPoi() throws IOException {
		rotateExcel("1691000a", "C:\\Users\\arrizaqu\\Desktop\\ExcelTemplate.xlsx", "C:\\Users\\arrizaqu\\Desktop\\ExcelTemplateB.xlsx", 5);
	}
	
	public static void rotateExcel(String formCode, String path, String dest, int maxColumn) throws IOException {
		FileInputStream file = new FileInputStream(new File(path));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(formCode);
		DataFormatter dataFormatter = new DataFormatter();
		
		//copy the last list
		List<MetaCell> copyData = new ArrayList();
		for (Row row: sheet) {
			MetaCell mc = new MetaCell();
			mc.setValue(dataFormatter.formatCellValue(row.getCell(maxColumn)));
			if(row.getCell(maxColumn) != null) {
				mc.setCellStyle(row.getCell(maxColumn).getCellStyle());
			}
			copyData.add(mc);
		}
	
		//shift 1 column
		sheet.shiftColumns(maxColumn-1, maxColumn-1, 1);
		
//		//paste data list
		int i = 0;
		for (Row row: sheet) {
			Cell cell = row.createCell(maxColumn-1);
			cell.setCellValue(copyData.get(i).getValue());
			cell.setCellStyle(copyData.get(i).getCellStyle());
			i++;
		}
		
        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(dest));
        workbook.write(out);
        
        //release
        workbook.close();
        out.close();		
        file.close();
	}
	
	public static void doubleRoutateExcel (String formCode, String path, String dest, int maxColumn) throws IOException {
		FileInputStream file = new FileInputStream(new File(path));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(formCode);
		DataFormatter dataFormatter = new DataFormatter();
		
		//copy the last list
		List<MetaCell> copyData = new ArrayList();
		for (Row row: sheet) {
			MetaCell mc = new MetaCell();
			mc.setValue(dataFormatter.formatCellValue(row.getCell(maxColumn)));
			if(row.getCell(maxColumn) != null) {
				mc.setCellStyle(row.getCell(maxColumn).getCellStyle());
			}
			copyData.add(mc);
		}
	}
}

class MetaCell{
	
	String value;
	CellStyle cellStyle;
	String cellFormula;
	
	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	public String getCellFormula() {
		return cellFormula;
	}

	public void setCellFormula(String cellFormula) {
		this.cellFormula = cellFormula;
	}
	
	
}
