package com.arrizaqu.demoexcelpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.util.FileSystemUtils;

import com.arrizaqu.demoexcelpoi.entity.MetaCell;
import com.arrizaqu.demoexcelpoi.entity.MetaCellMerge;

@SpringBootTest
public class TestHTML {

	@Test
	public void test() throws IOException {
		String path = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\1621100_example.html";
		String dest = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\1621100_result.html";
		
		Map<String, String> mm = new HashMap();
		File input = new File(path);
		Document doc = Jsoup.parse(input, "UTF-8");
		
		buildRotateHTML(doc);
		
		System.out.println(doc.toString());
		//output
		FileOutputStream out = new FileOutputStream(new File(dest));
		String text = doc.toString();
		byte[] mybytes = text.getBytes();
		
		out.write(mybytes);
	}
	
	@Test
	public void testMerger() throws IOException {
		String path = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\1621100_double.html";
		String dest = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\1621100_double_result.html";
		
		Map<String, String> mm = new HashMap();
		File input = new File(path);
		Document doc = Jsoup.parse(input, "UTF-8");
		
		//merge html
		reBuildHTMLMerger(doc);
		//rotatehtml
		buildRotateHTML(doc);
		
		//output
		FileOutputStream out = new FileOutputStream(new File(dest));
		String text = doc.toString();
		byte[] mybytes = text.getBytes();
		
		out.write(mybytes);
	}
	
	public void reBuildHTMLMerger(Document doc) {
		// TODO Auto-generated method stub
		Elements tRow = doc.select("table").get(0).select("tr");
		int rsType = checkRowSpanType(tRow);
		/*
		 * 2. table is single rowspan or double rowspan
		 * 1. is table will merged or not
		 * */
		
		int isMerge = checkRowSpanMerger(tRow);
		System.out.println("isMerge : "+ isMerge);
		if(isMerge != 0) {
			int startRow = getStartRow(tRow);
			int maxColumn = tRow.get(startRow).select("td").size()-1;
			int itemNum = Integer.parseInt(tRow.get(startRow).select("td").get(0).attr("rowspan"));
			int parentNum = (tRow.size()-startRow)/itemNum;
			
			//jika merger untuk single rowspan
			int startA = 0;
			if(isMerge == 1) {
				//mencari posisi column headerLeft tpi bukan di column 1 - 2
				Elements columns = tRow.get(startRow).select("td");
				for (int i = 0; i < columns.size(); i++) {
					if(columns.get(i).hasClass("rowHeaderLeft") && i >= 2) {
						startA = i;
					}
				}
				
				System.out.println("START A : "+ startA);
				//delete row
				int aRow = startRow;
				for (int i = 0; i < parentNum; i++) {
					//delete parentRow
					tRow.get(aRow).select("td").get(startA-3).remove();
					tRow.get(aRow).select("td").get(startA-3).remove();
					tRow.get(aRow).select("td").get(startA-3).remove();
					tRow.get(aRow).select("td").get(startA-3).remove();
					
					//delete itemRow
					for (int j = 1; j < itemNum; j++) {
						tRow.get(aRow+j).select("td").get(startA-4).remove();
						tRow.get(aRow+j).select("td").get(startA-4).remove();
					}
					aRow = aRow + itemNum;
				}
				
				//setup / merge column lable
				tRow.get(1).select("td").get(startA-4).remove();
				tRow.get(1).select("td").get(startA-4).remove();
				
			} else if(isMerge == 2) { //merge untuk double rowspan
				
			}
		}
		
	}
	
	public static int checkRowSpanMerger(Elements tRow) {
		
		/*
		 * 0 => not need to merger
		 * 1 => merger by type 1
		 * 2 => merger by type 2
		 * */
		
		int tableType = 0;
		int startRow = getStartRow(tRow);

		Elements columns = tRow.get(startRow).select("td");
		
		int headerNum = 0;
		for (int i = 0; i < columns.size(); i++) {
			if(columns.get(i).hasClass("rowHeaderLeft")) {
				headerNum++;
			}
		}
		
		if(headerNum == 6) {
			return 2;
		} else if(headerNum == 4) {
			return 1;
		}
		
		return 0;
	}
	
	public static void buildRotateHTML(Document doc) throws IOException {
		Elements tRow = doc.select("table").select("tr");
		int rsType = checkRowSpanType(tRow);
		
		if(rsType > 0) {
			if(rsType == 1) {
				rotateHTML(tRow);
			} else if(rsType == 2){
				rotateHTML_2(tRow);
			}
		} 
	}
	
	public static void rotateHTML(Elements tRow) throws IOException {
		
		//table definition
		int startRow = getStartRow(tRow);
		int maxColumn = getLengthColumn(tRow.get(startRow).select("td"));
		int itemNum = Integer.parseInt(tRow.get(startRow).select("td").get(0).attr("rowspan"));
		int parentNum = (tRow.size()-startRow)/itemNum;

		System.out.println("startRow : "+ startRow);
		System.out.println("maxColumn : "+ maxColumn);
		System.out.println("itemNum :"+ itemNum);
		System.out.println("parentNum : "+ parentNum);
		
		//setup variable 
		List<MetaCellMerge> cellRows = new ArrayList<>();
		int fr = startRow;
		for (int i = 0; i < parentNum; i++) {
			MetaCellMerge mcm = new MetaCellMerge(fr, 0, maxColumn-1, maxColumn-1);
			cellRows.add(mcm);
			fr = fr + itemNum;
		}
//		
//		//paste of prosess mirroring
		for(MetaCellMerge cm: cellRows) {
			String temp = tRow.get(cm.getFirstRow()).select("td").get(cm.getFirstCol()).toString();
			tRow.get(cm.getFirstRow()).select("td").get(cm.getFirstCol()).remove();
			tRow.get(cm.getFirstRow()).select("td").get(cm.getFirstCol()).after(temp);
		}
	}
	
	public static int getLengthColumn(Elements allColumns) {
		boolean isCount = true;
		int rs = -1;
		//Elements allColumns = ;
		for (int i = 0; i < allColumns.size(); i++) {
			
			//count only rowHeaderRight
			if(allColumns.get(i).hasClass("rowHeaderRight")) {
				isCount = false;
				rs++;
			}
			
			//count only field before rowHeaderRight
			if(isCount) {
				rs++;
			}
		}
		return rs;
	}
	
	public static int getLengthColumnMerger(Elements allColumns) {
		boolean isCount = true;
		int rs = -1;
		//Elements allColumns = ;
		for (int i = 0; i < allColumns.size(); i++) {
			
			//count only rowHeaderRight
			if(allColumns.get(i).hasClass("rowHeaderRight")) {
				isCount = false;
				rs++;
			}
			
			//count only field before rowHeaderRight
			if(isCount) {
				rs++;
			}
		}
		return rs;
	}
	
	public static int getStartRow(Elements tRow) {
		for (int i = 0; i < tRow.size(); i++) {
			if(tRow.get(i).select("td").get(0).hasClass("rowHeaderLeft")) {
				return i;
			}
		}
		
		return 0;
	}
	
	public static void rotateHTML_2(Elements tRow) {
		int parentNum = 0;
		int maxColumn = 0;
		
		//search getStartRow
		int startRow = getStartRow(tRow);
		//search parentNum
		String parentLength = tRow.get(startRow).select("td").get(0).attr("rowspan");
		parentNum = Integer.parseInt(parentLength);
		//search maxColumn
		maxColumn = getLengthColumn(tRow.get(startRow).select("td"));
		
		//search subParent
		int maxSp = 0;
		int op=0;
		List<Integer> headerSubParent = new ArrayList();
		for (int i = 0; i < tRow.size(); i++) {
			//first loop for default value
			if(i == startRow) {
				Element c = tRow.get(i).select("td").get(1);
				maxSp = Integer.parseInt(c.attr("rowspan").toString());			
				op = maxSp;
				headerSubParent.add(i);
			}
			
			if(i >= startRow && i < (startRow+parentNum)) {
				if(op >= 0) {
					if(op == 0) {
						op = Integer.parseInt(tRow.get(i).select("td").get(0).attr("rowspan").toString());
						headerSubParent.add(i);
					}
					op--;
				}
			}
		}
		
		//count headerLeftParentNum
		List<Integer> indexRe = new ArrayList();
		int lp = startRow;
		for(int i = startRow; i < tRow.size(); i++) {
			if(tRow.get(i).select("td").get(0).attr("rowspan").equals(String.valueOf(parentNum))) {
				indexRe.add(lp);
				lp = lp + parentNum;
			}
		}
		
		//remove base on parent num
		for (int i = 0; i < indexRe.size(); i++) {
			String copyParent = tRow.get(indexRe.get(i)).select("td").get(maxColumn-2).toString();
			tRow.get(indexRe.get(i)).select("td").get(maxColumn-2).remove();
			
			//paste after maxColumn
			tRow.get(indexRe.get(i)).select("td").get(maxColumn-1).after(copyParent);
		}
		
//		//get rowspan for subParent
		List<Integer> subParentNum = new ArrayList();
		int lpk = 0;
		int blockNum = (tRow.size()-startRow)/parentNum;
		for (int q = 0; q < blockNum; q++) {
			for (int i = 0; i < headerSubParent.size(); i++) {
				if(i == 0) {
					Element cbt = tRow.get(headerSubParent.get(i)+lpk).select("td").get(maxColumn-2);
					cbt.remove();
					tRow.get(headerSubParent.get(i)+lpk).select("td").get(maxColumn-2).after(cbt.toString());
				}else {
					Element cbt =  tRow.get(headerSubParent.get(i)+lpk).select("td").get(maxColumn - 3);
					cbt.remove();
					tRow.get(headerSubParent.get(i)+lpk).select("td").get(maxColumn - 3).after(cbt.toString());
				}
			}
			lpk = lpk + parentNum;
		}
	}
	
	public static int checkRowSpanType(Elements tRow) {
		int tableType = 0;
		int startRow = getStartRow(tRow);

		Elements columns = tRow.get(startRow).select("td");
		
		int headerNum = 0;
		for (int i = 0; i < columns.size(); i++) {
			if(i <= 2) {
				if(columns.get(i).hasClass("rowHeaderLeft")) {
					headerNum++;
				}
			}
		}
		
		if(headerNum == 2) {
			return 1;
		} else if(headerNum == 3) {
			return 2;
		} else {
			return 0;
		}
	}
}
