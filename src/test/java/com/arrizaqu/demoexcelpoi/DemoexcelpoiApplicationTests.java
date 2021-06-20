package com.arrizaqu.demoexcelpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
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
		String path = "C:\\Users\\arrizaqu\\Desktop\\ExcelTemplate.xlsx";
		String dest = "C:\\Users\\arrizaqu\\Desktop\\ExcelTemplateC.xlsx";
		buildReConstructExcel(path, dest);
	}
	
	
	@Test
	void testPoi2() throws IOException {
		String path = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\ExcelTemplate.xlsx";
		String destination = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\ExcelTemplateC.xlsx";
		buildReConstructExcel_2(path, destination);
	}
	
	private void buildReConstructExcel(String path, String dest) throws IOException {
		FileInputStream file = new FileInputStream(new File(path));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		//data source formcode
		List<ExcelFormCode> efcs = new ArrayList();
		efcs.add(new ExcelFormCode("1697000",7, 7, 24, null, 13));
		efcs.add(new ExcelFormCode("1694000a", 5, 7, 24, null, 13));
		efcs.add(new ExcelFormCode("1693000", 5, 7, 24, null, 13));
		efcs.add(new ExcelFormCode("1692000", 7, 7, 24, null, 13));
		efcs.add(new ExcelFormCode("1691000a", 5, 7, 24, null, 13));
		efcs.add(new ExcelFormCode("1640300", 5, 7, 2, null, 12));
		efcs.add(new ExcelFormCode("1640200", 4, 7, 3, null, 31));
		efcs.add(new ExcelFormCode("1640100", 4, 7, 2, null, 13));
		efcs.add(new ExcelFormCode("1621100", 8, 7, 12, null, 13));
		efcs.add(new ExcelFormCode("1621000a", 6, 7, 12, null, 13));
		efcs.add(new ExcelFormCode("1620300", 7, 7, 2, null, 12));
		efcs.add(new ExcelFormCode("1620200", 6, 7, 3, null, 31));
		efcs.add(new ExcelFormCode("1620100", 6, 7, 2, null, 13));
		efcs.add(new ExcelFormCode("1612000", 12, 7, 3, null, 8));
		efcs.add(new ExcelFormCode("1611000", 12, 7, 3, null, 38));
		
		//execute re-constract
		for(ExcelFormCode efc: efcs) {
			rotateMergeEfectPoi3(workbook, efc);
		}
		
		//Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(dest));
        workbook.write(out);
        
        //release
        file.close();
        out.close();
	}
	
	private void buildReConstructExcel_2(String path, String dest) throws IOException {
		FileInputStream file = new FileInputStream(new File(path));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		//data source formcode
		List<ExcelFormCode> efcs = new ArrayList();
		List<ExcelFormCode> efcDoubleMerge = new ArrayList();
		
		efcDoubleMerge.add(new ExcelFormCode("4625100",6, 5, 3, "3,3,3", 0));
		efcDoubleMerge.add(new ExcelFormCode("4623100",6, 5, 2, "3,3,3", 0));
		efcDoubleMerge.add(new ExcelFormCode("4622100",6, 5, 2, "3,3,3", 0));
		efcDoubleMerge.add(new ExcelFormCode("4621100",6, 5, 2, "3,3,3", 0));
		efcDoubleMerge.add(new ExcelFormCode("4611100a",6, 7, 3, "3,3,3", 0));
		
		efcs.add(new ExcelFormCode("4695000", 5, 7, 24, null, 13));
		efcs.add(new ExcelFormCode("4626100", 4, 5, 3, null, 3));
		efcs.add(new ExcelFormCode("4624100", 4, 5, 6, null, 3));
		efcs.add(new ExcelFormCode("4613100a", 9, 9, 3, null, 12));
		efcs.add(new ExcelFormCode("4613200a", 9, 9, 3, null, 12));
		efcs.add(new ExcelFormCode("4612200a", 9, 9, 3, null, 9));
		efcs.add(new ExcelFormCode("4612100a", 9, 9, 3, null, 9));
		efcs.add(new ExcelFormCode("4612000", 12, 7, 3, null, 8));
		efcs.add(new ExcelFormCode("4611200a", 4, 7, 3, null, 3));
		efcs.add(new ExcelFormCode("4611000", 12, 7, 3, null, 26));

		//execute re-constract double rowspan / merge
		for(ExcelFormCode drs: efcDoubleMerge) {
			rotateExcelDoubleMergeEfectPoi3(workbook, drs);
		}
		
		//execute re-constract singgle rowspan / merge
		for(ExcelFormCode ds: efcs) {
			rotateMergeEfectPoi3(workbook, ds);
		}
		
		//Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(dest));
        workbook.write(out);
        
        //release
        file.close();
        out.close();
	}

	/*
	 * 1. clear merge of moving column
	 * 2. copy sub parent and removing cell
	 * 3. paste or write data cell to destination column
	 * 4. adding merging new final colum of paste
	 * */
	public static void rotateMergeEfectPoi3(XSSFWorkbook workbook, ExcelFormCode efc) throws IOException {
		XSSFSheet sheet = workbook.getSheet(efc.getFormCode());
		DataFormatter dataFormatter = new DataFormatter();
		int columnParent = efc.getMaxColumn();
		
		//clear old merge column
		List<MetaCellMerge> cm1 = new ArrayList();
		
		int fr = efc.getStartRow();
		int lr = 0;
		for (int i = 0; i < efc.getParentNum(); i++) {
			lr = fr + efc.getItemNum() - 1;
			//System.out.println("pos : "+i+", fr : "+ fr + ", lr: "+ lr + ", colum : "+ (efc.getMaxColumn() - 1));
			cm1.add(new MetaCellMerge(fr, lr, efc.getMaxColumn()-1, efc.getMaxColumn()-1));
			fr = fr + efc.getItemNum();
		}	
		
		for (MetaCellMerge mc : cm1) {
    		for(int i=0; i < sheet.getNumMergedRegions(); ++i) {
	            CellRangeAddress range = sheet.getMergedRegion(i);    
	            if ( range.getFirstRow() == mc.getFirstRow() &&
	        			range.getFirstColumn() == mc.getFirstCol() &&
	             		range.getLastColumn() == mc.getLastCol() &&
	             		range.getLastRow() == mc.getLastRow()
	             		){
	                 sheet.removeMergedRegion(i);
	        	 }
	        }
		}
		
		//copy
		List<MetaCell> parentCell = new ArrayList();
		List<MetaCell> itemCell = new ArrayList();
		
		for (Row row: sheet) {
        	if(row.getRowNum() >= efc.getStartRow()) {
        		MetaCell mc = new MetaCell();
        		MetaCell mc0 = new MetaCell();
        		
        		Cell cell = row.getCell(efc.getMaxColumn()-1);
        		Cell cell0 = row.getCell(efc.getMaxColumn());
        		
        		String name = dataFormatter.formatCellValue(cell);
            	String name0 = dataFormatter.formatCellValue(cell0);
        		
            	mc.setValue(name);
            	mc0.setValue(name0);
            	
            	if(cell != null) {
            		mc.setCellStyle(cell.getCellStyle());
            	}
            	if(cell0 != null) {
            		mc0.setCellStyle(cell0.getCellStyle());
            	}
            	
            	mc.setRowNum(row.getRowNum());
            	mc0.setRowNum(row.getRowNum());
            	
            	mc.setSourceColumnNum(efc.getMaxColumn()-1);
            	mc0.setSourceColumnNum(efc.getMaxColumn());
            	
            	mc.setDestColumnNum(efc.getMaxColumn());
            	mc0.setDestColumnNum(efc.getMaxColumn()-1);
            	
    			parentCell.add(mc);
    			itemCell.add(mc0);
    			
    			//clear column
    			if(cell != null) {
    				row.removeCell(cell);
    			}
    			
    			if(cell0 != null) {
    				row.removeCell(cell0);
    			}
        	}
		}
		
		List<List<MetaCell>> dataCell = new ArrayList();
		dataCell.add(parentCell);
		dataCell.add(itemCell);
		
		//paste or write data cell to destination column
        for(List<MetaCell> mck: dataCell) {
        	for(MetaCell mc: mck) {
            	Row row = sheet.getRow(mc.getRowNum());
            	Cell cell = row.createCell(mc.getDestColumnNum());
            	cell.setCellValue(mc.getValue());
            	cell.setCellStyle(mc.getCellStyle());
            }
        }
		
		//add merge
		for (MetaCellMerge mc : cm1) {
    		sheet.addMergedRegion(new CellRangeAddress(mc.getFirstRow(), mc.getLastRow(), (mc.getFirstCol()+1), (mc.getLastCol()+1)));
		}

	}
	
	public static void rotateExcelDoubleMergeEfectPoi3(XSSFWorkbook workbook, ExcelFormCode efc) throws IOException {
		XSSFSheet sheet = workbook.getSheet(efc.getFormCode());
		DataFormatter dataFormatter = new DataFormatter();
		int columnDest = efc.getMaxColumn()+1;
		int columnParent = columnDest+1;
		
		String[] pN = efc.getSubParentNum().split(",");
		
		int totalPn = 0;
		for (int i = 0; i < pN.length; i++) {
			totalPn = totalPn + Integer.parseInt(pN[i]);
		}
		
		int longItem = totalPn*efc.getParentNum();
		//clear merge
		/*
		 * 1. clear merge if exist in the right side
		 * 2. clear merge of moving column
		 * 3. copy sub parent and removing cell
		 * 4. paste or write data cell to destination column
		 * 5. adding merging new final colum of paste
		 * */
		
//		for(int p = startRow; p < startRow+longItem; p++ ) {
//			for(int i=0; i < sheet.getNumMergedRegions(); ++i) {
//	            CellRangeAddress range = sheet.getMergedRegion(i);    
//            	if ( range.getFirstRow() == p &&
//	        			range.getFirstColumn() == columnDest &&
//	             		range.getLastColumn() == columnParent &&
//	             		range.getLastRow() == p
//	             		){
//	                 sheet.removeMergedRegion(i);
//	        	 }
//	            
//		    }
//		}
        
        //clear merge of moving column 
        List<List<MetaCellMerge>> cms = new ArrayList();
        int fr = efc.getStartRow();
        int fr0 = efc.getStartRow();
        List<MetaCellMerge> cm1 = new ArrayList();
        List<MetaCellMerge> cmParent = new ArrayList();
        
  	    for (int i = 0; i < efc.getParentNum(); i++) {
  			for (int j = 0; j < pN.length; j++) {
  				int z = Integer.parseInt(pN[j]);
  				int lr = (fr+(z-1));
  				
  				//add to remove merge for sub parent
  				cm1.add(new MetaCellMerge(fr, lr, efc.getMaxColumn()-1, efc.getMaxColumn()-1));
  				fr = fr + z;
  			}
  			
  			cmParent.add(new MetaCellMerge(fr0, fr0+(totalPn-1), efc.getMaxColumn()-2, efc.getMaxColumn()-2));
  			fr0 = fr0+totalPn;
  		}
  	    
  	    cms.add(cm1);
	    cms.add(cmParent);
	    
	    for(List<MetaCellMerge> mad: cms) {
	    	for (MetaCellMerge mc : mad) {
	    		for(int i=0; i < sheet.getNumMergedRegions(); ++i) {
		            CellRangeAddress range = sheet.getMergedRegion(i);    
		            if ( range.getFirstRow() == mc.getFirstRow() &&
		        			range.getFirstColumn() == mc.getFirstCol() &&
		             		range.getLastColumn() == mc.getLastCol() &&
		             		range.getLastRow() == mc.getLastRow()
		             		){
		                 sheet.removeMergedRegion(i);
		        	 }
		        }
			}
        }
	    
		//copy sub parent and removing cell
		List<MetaCell> subParenCell = new ArrayList();
		List<MetaCell> parentCell = new ArrayList();
		
        for (Row row: sheet) {
        	if(row.getRowNum() >= efc.getStartRow()) {
        		MetaCell mc = new MetaCell();
        		MetaCell mcP = new MetaCell();
        		
            	Cell cell = row.getCell(efc.getMaxColumn());
            	Cell cellP = row.getCell(efc.getMaxColumn() - 2);
            	
            	String name = dataFormatter.formatCellValue(cell);
            	String nameParent = dataFormatter.formatCellValue(cellP);
            	
            	mc.setValue(name);
            	mcP.setValue(nameParent);
            	
            	mc.setCellStyle(cell.getCellStyle());
            	mcP.setCellStyle(cellP.getCellStyle());
            	
            	mc.setRowNum(row.getRowNum());
            	mcP.setRowNum(row.getRowNum());
            	
            	mc.setSourceColumnNum(efc.getMaxColumn());
            	mcP.setSourceColumnNum(efc.getMaxColumn() - 2);
            	
            	mc.setDestColumnNum(efc.getMaxColumn()-2);
            	mcP.setDestColumnNum(efc.getMaxColumn());
            	
    			subParenCell.add(mc);
    			parentCell.add(mcP);
    			
    			//clear column
    			row.removeCell(cell);
    			row.removeCell(cellP);
        	}
		}
        
        List<List<MetaCell>> dataCell = new ArrayList();
        dataCell.add(subParenCell);
        dataCell.add(parentCell);
        
        //paste or write data cell to destination column
        for(List<MetaCell> mck: dataCell) {
        	for(MetaCell mc: mck) {
            	Row row = sheet.getRow(mc.getRowNum());
            	Cell cell = row.createCell(mc.getDestColumnNum());
            	cell.setCellValue(mc.getValue());
            	cell.setCellStyle(mc.getCellStyle());
            }
        }
        
      //add merge new column 
	    int pq = 0;
	    for(List<MetaCellMerge> mad: cms) {
	    	for (MetaCellMerge mc : mad) {
	    		sheet.addMergedRegion(new CellRangeAddress(mc.getFirstRow(), mc.getLastRow(), (mc.getFirstCol()+pq), (mc.getLastCol()+pq)));
			}
	    	pq = pq + 2;
        }
	}
}

class MetaCellMerge{
	
	public MetaCellMerge(int firstRow, int lastRow, int firstCol, int lastCol) {
		super();
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
	}
	private int firstRow;
	private int lastRow;
	private int firstCol;
	private int lastCol;
	public int getFirstRow() {
		return firstRow;
	}
	public void setFirstRow(int firstRow) {
		this.firstRow = firstRow;
	}
	public int getLastRow() {
		return lastRow;
	}
	public void setLastRow(int lastRow) {
		this.lastRow = lastRow;
	}
	public int getFirstCol() {
		return firstCol;
	}
	public void setFirstCol(int firstCol) {
		this.firstCol = firstCol;
	}
	public int getLastCol() {
		return lastCol;
	}
	public void setLastCol(int lastCol) {
		this.lastCol = lastCol;
	}
}

class MetaCell{
	
	String value;
	CellStyle cellStyle;
	String cellFormula;
	int rowNum;
	int sourceColumnNum;
	int destColumnNum;
	
	public int getSourceColumnNum() {
		return sourceColumnNum;
	}

	public void setSourceColumnNum(int sourceColumnNum) {
		this.sourceColumnNum = sourceColumnNum;
	}

	public int getDestColumnNum() {
		return destColumnNum;
	}

	public void setDestColumnNum(int destColumnNum) {
		this.destColumnNum = destColumnNum;
	}

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

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

class ExcelFormCode{
	private int maxColumn;
	private String formCode;
	private int startRow;
	private int parentNum;
	private String subParentNum;
	private int itemNum;
	
	public ExcelFormCode() {}
	
	public ExcelFormCode(String formCode,int maxColumn, int startRow, int parentNum,
			String subParentNum, int itemNum) {
		super();
	
		this.maxColumn = maxColumn;
		this.formCode = formCode;
		this.startRow = startRow;
		this.parentNum = parentNum;
		this.subParentNum = subParentNum;
		this.itemNum = itemNum;
	}
	
	public int getItemNum() {
		return itemNum;
	}

	public void setItemNum(int itemNum) {
		this.itemNum = itemNum;
	}

	public int getMaxColumn() {
		return maxColumn;
	}
	public void setMaxColumn(int maxColumn) {
		this.maxColumn = maxColumn;
	}
	public String getFormCode() {
		return formCode;
	}
	public void setFormCode(String formCode) {
		this.formCode = formCode;
	}
	public int getStartRow() {
		return startRow;
	}
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}
	public int getParentNum() {
		return parentNum;
	}
	public void setParentNum(int parentNum) {
		this.parentNum = parentNum;
	}
	public String getSubParentNum() {
		return subParentNum;
	}
	public void setSubParentNum(String subParentNum) {
		this.subParentNum = subParentNum;
	}
	
}
