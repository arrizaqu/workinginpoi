package com.arrizaqu.demoexcelpoi.entity;

import org.apache.poi.ss.usermodel.CellStyle;

public class MetaCell{
	
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