package com.arrizaqu.demoexcelpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

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
		
		//FileInputStream fileHTML = new FileInputStream(new File(path));
		
		File input = new File(path);
		Document doc = Jsoup.parse(input, "UTF-8");
		
		Element table = doc.select("table").get(0);
		
		//formula
		int maxColumn = 8;
		int parentNum = 12;
		int startRow = 2;
		int itemNum = 13;
		
		//ok
		Elements row = table.select("tr");
		//setup
		List<MetaCellMerge> cellRows = new ArrayList<>();
		List<Integer> lineDetected = new ArrayList<>();
		
		int fr = startRow+1;
		for (int i = 0; i < parentNum; i++) {
			//System.out.println("fr : "+ fr);
			MetaCellMerge mcm = new MetaCellMerge(fr, 0, 0, 0);
			cellRows.add(mcm);
			//System.out.println("fr:"+fr+", i:"+i);	
			lineDetected.add(fr);
			fr = fr + itemNum;
		}
		
		
		//extract
		Elements td = row.get(3).select("td");
		td.get(7).remove();
		
		System.out.println("value : "+ row.get(3).select("td").get(7).text());
		System.out.println("value : "+ row.get(4).select("td").get(8).text());
		
		for(MetaCellMerge cm: cellRows) {
			//System.out.println("cm : "+ cm.getFirstRow());
			
			//remove parentColumn
			//row.get(cm.getFirstRow()).select("td").get(maxColumn-1).remove();
			
			//System.out.println(""+ row.get(cm.getFirstRow()).select("td").get(maxColumn-1).text());
		}
		
		
		
		
		
		//System.out.println("DOM : "+ table.toString());
		FileOutputStream out = new FileOutputStream(new File(dest));
		final File f = new File("filename.html");
        //FileSystemUtils.writeStringToFile(f, doc.outerHtml(), "UTF-8");
		
	}
}
