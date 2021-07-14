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
		String path = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\new\\1694000a.html";
		String dest = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\new\\1694000a_result.html";
		
		Map<String, String> mm = new HashMap();
		File input = new File(path);
		Document doc = Jsoup.parse(input, "UTF-8");
		
		//merge html
		reBuildHTMLMerger2(doc);
		//rotatehtml
		//buildRotateHTML(doc);
		
		//output
		FileOutputStream out = new FileOutputStream(new File(dest));
		String text = doc.toString();
		byte[] mybytes = text.getBytes();
		
		out.write(mybytes);
	}
	
	@Test //example merger for double rowspan
	public void testMergerDoubleRowspan() throws IOException {
		String path = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\8694000a.html";
		String dest = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\8694000a_result.html";
		
		File input = new File(path);
		Document doc = Jsoup.parse(input, "UTF-8");
		
		//merge html
		//reBuildHTMLMerger(doc);
		//rotatehtml
		//buildRotateHTML(doc);
		
		//output
		FileOutputStream out = new FileOutputStream(new File(dest));
		String text = doc.toString();
		byte[] mybytes = text.getBytes();
		
		out.write(mybytes);
	}
	
	public int getHeaderColumnNum(Element tBle) {
		Elements rows = tBle.select("tr");
		for (int i = 0; i < rows.size(); i++) {
			Elements columns = rows.get(i).select("td");
			for (int j = 0; j < columns.size(); j++) {
				if(columns.get(j).hasClass("dimensionalFormColumnHeader")) {
					return i;
				}
			}
		}
		
		return 0;
	}
	
	public void reBuildHTMLMerger2(Document doc) {
		Elements tables = doc.select("table.table01");
		
		Element tableLeft = tables.get(0);
		Element tableRight = tables.get(1);
		
		Elements tRow = tableLeft.select("tr");
		Elements tRowRight = tableRight.select("tr");
		
		int rsType = checkRowSpanType(tRow);
		
		//remove title
		Element tableDate = doc.select("table").get(0);
		int maxColumn = 0;
		if(rsType == 0) { // tanpa rowspan
			int startRow = getStartRow(tRow);
			maxColumn = tRow.get(startRow).select("td").size()-1;
			int itemNum = tRow.size()-startRow;
			int columnRightNum = (tRowRight.get(startRow).select("td").size()-2);
			
			//int parentNum = (tRow.size()-startRow)/itemNum;
			int cTableNum = getHeaderColumnNum(tableLeft);
			
			System.out.println("ctable num : "+ cTableNum);
			System.out.println("startRow : "+ startRow);
			System.out.println("maxColumn : "+ maxColumn);
			System.out.println("itemNum :"+ itemNum);
			System.out.println("columnRightNum : "+ columnRightNum);
			
			//copy column header to left
			for (int i = 0; i < tRow.size(); i++) {
				Elements columns = tRow.get(i).select("td");
				if(columns.get(1).hasClass("dimensionalFormColumnHeader")) {

					int a = 1;
					for (int j = 0; j < columnRightNum; j++) {
						Element el=null;
						if(tRow.get(i).select("td").get(0).hasAttr("rowspan")) {
							Element lastField = tRow.get(i).select("td").get(maxColumn-1+j);
							el = tRowRight.get(i).select("td").get(a);
							lastField.after(el.toString());
						} else {
							Element lastField = tRow.get(i).select("td").get(maxColumn-2+j);
							el = tRowRight.get(i).select("td").get(a-1);
							lastField.after(el.toString());
						}
						a++;
					}
				}	
			}
			
			for (int i = startRow; i < tRow.size(); i++) {
				Element lastField = tRow.get(i).select("td").get(maxColumn-1);
				
				//copy field from right table
				List<String> rightCopy = new ArrayList();
				int a = 0;
				for (int j = 0; j < columnRightNum; j++) {
					Element el = tRowRight.get(i).select("td").get(++a);
					rightCopy.add(el.toString());
				}
				
				//paste to left
				for (int j = 0; j < rightCopy.size(); j++) {
					lastField.after(rightCopy.get(j));
				}

			}
			
			//get maxColumn after merge
			int newMaxColumn = tableLeft.select("tr").get(startRow).select("td").size();
			
			//settup titleLeft & right
			tableLeft.select("tr").get(0).select("td").get(0).attr("colspan", String.valueOf(maxColumn));
			tableLeft.select("tr").get(0).select("td").get(1).attr("colspan", String.valueOf(newMaxColumn - maxColumn));
			
			//add row
			Elements tt = tableDate.getElementsByClass("raportado_title_date");
			String titlCP = "<td colspan='"+maxColumn+"'><p class='titleLeft'>"+tt.get(0)+"</p></td>";
			titlCP += "<td colspan='"+(newMaxColumn - maxColumn)+"'><p class='titleRight'>"+tt.get(1)+"</p></td>";
			tableLeft.select("tr").get(0).before(titlCP);
			//delete row tt
			tt.remove();
			
			//delete right table
			tableRight.remove();
		} else if(rsType == 1) {
			int startRow = getStartRow(tRow);
			maxColumn = tRow.get(startRow).select("td").size()-1;
			int itemNum = Integer.parseInt(tRow.get(startRow).select("td").get(0).attr("rowspan"));
			int columnRightNum = (tRowRight.get(startRow).select("td").size()-4); //=> rowspan
			
			int parentNum = (tRow.size()-startRow)/itemNum;
			int cTableNum = getHeaderColumnNum(tableLeft);
			
			System.out.println("ctable num : "+ cTableNum);
			System.out.println("startRow : "+ startRow);
			System.out.println("maxColumn : "+ maxColumn);
			System.out.println("itemNum :"+ itemNum);
			System.out.println("parentNum : "+ parentNum);
			System.out.println("columnRightNum : "+ columnRightNum);
			
			//copy header to left
			for (int i = 0; i < tRow.size(); i++) {
				Elements columns = tRow.get(i).select("td");
				if(columns.get(1).hasClass("dimensionalFormColumnHeader")) {

					int a = 1;
					for (int j = 0; j < columnRightNum; j++) {
						Element el=null;
						if(tRow.get(i).select("td").get(0).hasAttr("rowspan")) {
							Element lastField = tRow.get(i).select("td").get(maxColumn-3+j);
							el = tRowRight.get(i).select("td").get(a);					
							lastField.after(el.toString());
						} else {
							Element lastField = tRow.get(i).select("td").get(maxColumn-4+j);
							el = tRowRight.get(i).select("td").get(a-1);
							lastField.after(el.toString());
						}
						a++;
					}
				}	
			}
			
			//copy field content to left
			for (int i = startRow; i < tRow.size(); i++) {
				Element lastField = null;
				//cek if there is rowspan
				if(tRow.get(i).select("td").get(0).hasAttr("rowspan")) {
					lastField = tRow.get(i).select("td").get(maxColumn-2);
				} else {
					lastField = tRow.get(i).select("td").get(maxColumn-3);
				}
				
				//copy field from right table
				List<String> rightCopy = new ArrayList();
				int a = 1;
				for (int j = 0; j < columnRightNum; j++) {
					a++;
					if(tRowRight.get(i).select("td").get(0).hasAttr("rowspan")) {
						Element el = tRowRight.get(i).select("td").get(a);
						rightCopy.add(el.toString());
					} else {
						Element el = tRowRight.get(i).select("td").get(a-1);
						rightCopy.add(el.toString());
					}
				}
				
				//paste to left
				for (int j = 0; j < rightCopy.size(); j++) {
					lastField.after(rightCopy.get(j));
				}
			}
			
			//get maxColumn after merge
			int newMaxColumn = tableLeft.select("tr").get(startRow).select("td").size();
			
			//settup titleLeft & right
			tableLeft.select("tr").get(0).select("td").get(0).attr("colspan", String.valueOf(maxColumn));
			tableLeft.select("tr").get(0).select("td").get(1).attr("colspan", String.valueOf(newMaxColumn - maxColumn));
			
			//add row
			Elements tt = tableDate.getElementsByClass("raportado_title_date");
			String titlCP = "<td colspan='"+maxColumn+"'><p class='titleLeft'>"+tt.get(0)+"</p></td>";
			titlCP += "<td colspan='"+(newMaxColumn - maxColumn)+"'><p class='titleRight'>"+tt.get(1)+"</p></td>";
			tableLeft.select("tr").get(0).before(titlCP);
			//delete row tt
			tt.remove();
			
			//delete right table
			tableRight.remove();
		}
	}
	
	public void reBuildHTMLMerger(Document doc) {
		// TODO Auto-generated method stub
		
		Elements generalTable = doc.select("table.table01");
		
		System.out.println("general table : "+ generalTable.size());
		
		int ct = 0;
		
//		for (int i = 0; i < generalMerge.size(); i++) {
//			Element table = generalMerge.get(i);
//			if(table.hasClass("table01")) {
//				ct++;
//			}
//		}
		
		Element table = doc.select("table").get(0);
		Elements tRow = table.select("tr");
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
			int cTableNum = getHeaderColumnNum(table);
			
//			System.out.println("ctable num : "+ cTableNum);
//			System.out.println("startRow : "+ startRow);
//			System.out.println("maxColumn : "+ maxColumn);
//			System.out.println("itemNum :"+ itemNum);
//			System.out.println("parentNum : "+ parentNum);
			
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
				tRow.get(cTableNum).select("td").get(startA-4).remove();
				tRow.get(cTableNum).select("td").get(startA-4).remove();
				
			} else if(isMerge == 2) { //merge untuk double rowspan
				//mencari posisi column headerLeft tpi bukan di column 1 - 2
				Elements columns = tRow.get(startRow).select("td");
				for (int i = 0; i < columns.size(); i++) {
					if(columns.get(i).hasClass("rowHeaderLeft") && i >= 2) {
						startA = i;
					}
				}
				
				System.out.println("start A : "+ startA);
				
				//delete row parent
				int aRow = startRow;
				int add = 0;
				List<Integer> mergePoints = new ArrayList();
				for (int i = 0; i < parentNum; i++) {
					//delete parentRow
					tRow.get(aRow).select("td").get(startA-5).remove();
					tRow.get(aRow).select("td").get(startA-3).remove();
					tRow.get(aRow).select("td").get(startA-3).remove();
					//System.out.println("abo : "+ tRow.get(aRow).select("td").get(startA-5).toString());
					int defultV = Integer.parseInt(tRow.get(aRow).select("td").get(startA-5).attr("rowspan"));
					int dst = defultV+startRow;
					
					//delete itemRow
					for (int j = 0; j < defultV; j++) {
						if(j == 0) {
							//delete column sayap kanan (table kiri)
							tRow.get(aRow+j).select("td").get(startA-5).remove();
							tRow.get(aRow+j).select("td").get(startA-5).remove();
							tRow.get(aRow+j).select("td").get(startA-5).remove();
							
							//get rs for parent after
							int ab = 0;
							for (int j2 = 0; j2 < defultV-1; j2++) {
								mergePoints.add((dst+add)+ab);
								ab = ab + defultV;
							}
							add = add + itemNum;
						} else {
							//delete column sayap kanan (table kiri)
							tRow.get(aRow+j).select("td").get(startA-6).remove();
							tRow.get(aRow+j).select("td").get(startA-7).remove();
						}
						
					}
					aRow = aRow + itemNum;
				}
				
				for (int i = 0; i < mergePoints.size(); i++) {
					System.out.println("pos"+mergePoints.get(i)+"ws: "+ tRow.get(mergePoints.get(i)).select("td").get(0).toString());
					int defultV2 = Integer.parseInt(tRow.get(mergePoints.get(i)).select("td").get(0).attr("rowspan"));
					for (int j = 0; j < defultV2; j++) {
						if(j == 0) {
							//delete column sayap kanan (table kiri)
							tRow.get(mergePoints.get(i)+j).select("td").get(startA-6).remove();
							tRow.get(mergePoints.get(i)+j).select("td").get(startA-6).remove();
							tRow.get(mergePoints.get(i)+j).select("td").get(startA-6).remove();
							tRow.get(mergePoints.get(i)+j).select("td").get(startA-6).remove();
						} else {
							//delete column sayap kanan (table kiri)
							tRow.get(mergePoints.get(i)+j).select("td").get(startA-6).remove();
							tRow.get(mergePoints.get(i)+j).select("td").get(startA-7).remove();
						}
					}
				}
				
				if(cTableNum != 0) {
					tRow.get(cTableNum).select("td").get(startA-7).remove();
					tRow.get(cTableNum).select("td").get(startA-7).remove();
				}
//				int defultV2 = Integer.parseInt(tRow.get(9).select("td").get(0).attr("rowspan"));
//				for (int j = 0; j < defultV2; j++) {
//					if(j == 0) {
//						//delete column sayap kanan (table kiri)
//						tRow.get(9+j).select("td").get(startA-6).remove();
//						tRow.get(9+j).select("td").get(startA-6).remove();
//						tRow.get(9+j).select("td").get(startA-6).remove();
//						tRow.get(9+j).select("td").get(startA-6).remove();
//					} else {
//						//delete column sayap kanan (table kiri)
//						tRow.get(9+j).select("td").get(startA-6).remove();
//						tRow.get(9+j).select("td").get(startA-7).remove();
//					}
//				}
//				
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

		System.out.println("start row : "+ startRow);
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

//		System.out.println("startRow : "+ startRow);
//		System.out.println("maxColumn : "+ maxColumn);
//		System.out.println("itemNum :"+ itemNum);
//		System.out.println("parentNum : "+ parentNum);
		
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
				System.out.print("post : "+ i);
				System.out.println(", sss : "+ tRow.get(i).select("td").get(0).toString());
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
