package ae999;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

/**
 * Created by ae on 06.07.2016.
 *
 * Обработать таблицы в документе RTF
 * Для этого RTF переведем в DOCX
 * Обработаем DOCX
 * А затем переведем DOCX ообратно в RTF
 */
public class TableRTF extends TableDOCX {


    boolean parse(File iFile) {
        // return super.parse(inputFileName, outputFileName);
    	// сначала сделаем копию входного файла
    	boolean result=false;
        String fn, pf;
        fn = iFile.getAbsolutePath();
    	pf = fn + ".docx";
    	// преобразуем входной файл rtf в docx
    	if (RunWord.rtf2docx(fn, pf)) {
            //
            File f = new File(pf);
            //
            result = super.parse(f) ;
            //
            f.delete();
    	}
    	return result;
    } // end parse()
    
    
    // простой и удобный метод копирования файла в Java 7
    public static boolean copyFile2File(File source, File dest) 
    {
    	try {
        	Files.copy(source.toPath(), dest.toPath());
        	return true;
        } catch (IOException ex){
        	ex.printStackTrace();
        }
        return false;
    } // end copyFile2File()
    
    // простой и удобный метод копирования файла в Java 7
    public static void ex_copyFile2File(File source, File dest) throws IOException 
    {
    	Files.copy(source.toPath(), dest.toPath());        	
    } // end copyFile2File()
    
}
