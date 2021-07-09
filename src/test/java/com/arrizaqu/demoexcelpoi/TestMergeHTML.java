package com.arrizaqu.demoexcelpoi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
public class TestMergeHTML {

	@Test
	public void test() throws IOException {
		String path = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\1621100_double.html";
		String dest = "C:\\Users\\arrizaqu\\Desktop\\combinasi\\1621100_double_result.html";
		
		Map<String, String> mm = new HashMap();
		File input = new File(path);
		Document doc = Jsoup.parse(input, "UTF-8");
		
		//buildRotateHTML(doc);
		//reBuildHTMLMerger(doc);
		
		//System.out.println(doc.toString());
		//output
//		FileOutputStream out = new FileOutputStream(new File(dest));
//		String text = doc.toString();
//		byte[] mybytes = text.getBytes();
//		
//		out.write(mybytes);
	}

	
}
