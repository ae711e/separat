package ae999;

import java.io.File;
import java.io.FilenameFilter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

// класс для фильтрации списка файлов, выдаваемых File.list
public class MyFilter implements FilenameFilter {
  //String strMask;
  Pattern pat;

	public MyFilter(String maskName) {
		//strMask=maskName;
		pat=Pattern.compile(maskName);
	}
	
	@Override
	public boolean accept(File dir, String name) {
		// TODO Auto-generated method
		Matcher m = pat.matcher(name);
		return m.matches();
	}
	  
}
