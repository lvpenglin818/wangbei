package wangbei.version1;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class Test {

	public static void main(String[] args) throws Exception {
		FileWriter fw = new FileWriter(new File("C://Users//lvpenglin//Desktop//out.csv"));
		fw.write("ab,cf"+"\r\n");
		fw.write("ab,cf");
		fw.write("ab,cf");
		fw.flush();
	}
	public static void sopTest(Map mapC){
		Iterator<Map.Entry<String, List>> it = mapC.entrySet().iterator();
		while(it.hasNext()){
			Map.Entry<String, List> me = it.next();
			System.out.println("key is:"+ me.getKey().toString()+"value is:"+ me.getValue().toString());
		}
	}
}
