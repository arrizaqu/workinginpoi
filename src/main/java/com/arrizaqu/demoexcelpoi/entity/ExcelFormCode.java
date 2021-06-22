package com.arrizaqu.demoexcelpoi.entity;

public class ExcelFormCode{
	//maxcolummn
	private int maxColumn;
	//formCode
	private String formCode;
	//startRow
	private int startRow;
	//parentNum, list parent 
	private int parentNum;
	//just for double rowspan, if 1 rs => null, ex: 3,3,3
	private String subParentNum;
	//just for 1 rs, if double rs => 0, ex: 9 per block
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