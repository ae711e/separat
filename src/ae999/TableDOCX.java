package ae999;

/*
 * Created by Алексей on 03.07.2016.
 * (C) 2016, Aleksey Eremin
 *
 * обработка таблицы в файле  .docx
*/
// Modify:
// 2016-07-03 добавил очистку "странной ячейки" emptedStarngeCell()
//

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

// переделать таблицу со скриншотами в файле DOCX
public class TableDOCX {

    final protected static String s_separatorDir = System.getProperty("file.separator");
    final protected static String s_subdirJpg1 ="yYy";
    final protected static String s_subdirJpg2 ="x";
    
    boolean parse(File iFile)
    {
        final String strNEED_TEXT="Наименование файла скриншота.";   	// название требуемой колонки
        String str;
        int cn, cnt;
        int i, n;
        boolean result=false;
        String DiRi;		// название каталога с картинками
        FileInputStream fileInputStream;
        XWPFDocument docx;
        try {
            fileInputStream = new FileInputStream(iFile);
            // открываем файл и считываем его содержимое в объект XWPFDocument
            docx = new XWPFDocument(OPCPackage.open(fileInputStream));  // word 2007
            // Получим имя каталога входного файла
            DiRi=iFile.getParent();  // название каталога с картинками
            //
            fileInputStream.close();
        } catch(IOException | InvalidFormatException ex) {
            ex.printStackTrace();
            return false;
        }
        //
        Iterator<XWPFTable> tableIter;
        List<XWPFTableRow> rowList;
        List<XWPFTableCell> cellList;
        XWPFTable table;
        XWPFTableRow row;
        XWPFTableCell cell;
        XWPFTable mytable = null;  // требуемая таблица со списокм скриншотов
        int mycol=0; // колонка с именами файлов
        // пройдемся по таблицам документа в поисках нужной таблицы,
        // то есть такой, у которой в последнем столбце заголовка написано needText
        tableIter = docx.getTablesIterator();
        while(tableIter.hasNext()) {
            table = tableIter.next();		// очередная таблица
            //printTable(table);
            cellList=table.getRow(0).getTableCells(); // набор ячеек первой строки
            // ищем таблицу скриншотов
            cn=cellList.size()-1;  // индекс последней ячейки (столбца)
            // ищем таблицу со скриншотами
            if (cn == 4) {
                str=cellList.get(cn).getText();  // последняя колонка (там должны быть имена файлов)
                if( str.regionMatches(0, strNEED_TEXT, 0, strNEED_TEXT.length()) ) { // сравним с искомым текстом заоловка столбца
                    mytable=table;	// запомним таблицу
                    mycol=cn;	// запомним индекс столбца
                    break;
                }
            }
        } // end while
        //
        // вышли и проверим  - нашли ли таблицу?
        // считаем, если последняя колонка больше 0, значит нашли
        cnt = 0;
        if (mycol>1) {
            // создадим подкаталог, для файлов, которые есть в протоколе
            String jpgDir=DiRi+s_separatorDir+s_subdirJpg1;
            File wd1=new File(jpgDir);
            wd1.mkdir();
            
            //
            ////printTable(mytable);
            // пойдем по найденной таблице с конца и будем удалять строки у которых нет картинок
            rowList = mytable.getRows();        // список строк таблицы
            n = rowList.size();                                // кол-во строк таблицы
            // начнем проверять список файлов с конца таблицы,
            // и если файла нет - удалять строку
            for (i = n - 1; i > 0; i--) {
                row = rowList.get(i);                                        // взять строку с индексом i
                cellList = row.getTableCells();                // из строки получить список ячеек (колонки)
                str = cellList.get(mycol).getText();        // последняя колонка - имя файла с картинкой
                // если окончание (расширение) имени файла .jpg
                if (str.endsWith(".jpg")) {  // str.regionMatches(l-4, ".jpg",0,4) l=str.length(); // длина строки
                    // получили имя файла
                    iFile = new File(DiRi, str);
                    // проверим есть ли такой файл?
                    if (iFile.exists()) {
                        // файла нет
                        if (MoveIt(iFile, s_subdirJpg1)) {                // переместим файл
                            cnt++;
                        }
                        ////Log("Not exist " + cnt + ") " + str);
                    }  // if (!file.exists())
                } // end if jpg
                //
            }
            //
            // переместим оставшиеся файлы во временный каталог
            // создадим подкаталог, для файлов, которые не входят в протокол
            String outDir=DiRi+s_separatorDir+s_subdirJpg2;
            File wd2=new File(outDir);
            wd2.mkdir();
            int a;
            String maskExt=".*jpg$"; // ищем только jpg файлы
            MoveFiles(DiRi, outDir, maskExt);
            // переместим из первого каталога в основной
            MoveFiles(jpgDir, DiRi, maskExt);
            // удалим этот каталог
            wd1.delete();
            //
        }
        Log("Move it: " + cnt);
        //        
        try {
            docx.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        //
        return (cnt>0)? true: false;
    } // parse()

    // распечатать таблицу
    private static void printTable(XWPFTable table)
    {
        List<XWPFTableRow> rowList;
        Iterator<XWPFTableRow> rowIter;
        List<XWPFTableCell> cellList;
        Iterator<XWPFTableCell> cellIter;
        XWPFTableRow row;
        XWPFTableCell cell;
        System.out.println("[Begin table]");
        rowList = table.getRows();
        rowIter = rowList.iterator();
        while(rowIter.hasNext()) {
            row = rowIter.next();
            cellList = row.getTableCells();
            cellIter = cellList.iterator();
            while(cellIter.hasNext()) {
                cell = cellIter.next();
                System.out.print("|" + cell.getText());
            }
            System.out.println("|");
        }
        System.out.println("[End table]");
    }


    public static void Log(String str)
    {
        System.out.println(str);
    }

    // если убил - возвращает TRUE
    static boolean MoveIt(File ifile, String subdir)
    {
        String path=ifile.getParent();
        String name=ifile.getName();
        String outname;
        outname=path + s_separatorDir + subdir + s_separatorDir + name;
        File fn=new File(outname);
        if(!ifile.renameTo(fn)) {
           Log("?-ERROR-File don't move to: " + outname);
        }
        //
        Log(outname);
        //
        return true;
    }

    // переместить файлы  их одной директории (dir_src) в другую (dir_dst) с заданным расширением (ext)
    static int MoveFiles(String dir_src, String dir_dst, String ext)
    {
      int cnt=0;
        // проверим директорию - есть она и является ли директорией?
        File dirSrc=new File(dir_src);
        if(!dirSrc.exists() || !dirSrc.isDirectory()) {
            Log("?-ERROR-Not found directory: " + dir_src);
            return 0;
        }
        File dirDst=new File(dir_dst);
        if(!dirDst.exists() || !dirDst.isDirectory()) {
            Log("?-ERROR-Not found directory: " + dir_dst);
            return 0;
        }
    
        // получим имя директории, где файлы лежат
        String nameSrc;
        nameSrc=dirSrc.getAbsolutePath() + s_separatorDir;  // добавим разделитель
        // получим имя директории, куда файлы класть
        String nameDst;
        nameDst=dirDst.getAbsolutePath() + s_separatorDir;  // добавим разделитель
        
        // задается как Regexp выражение,
        String[] list=dirSrc.list(new MyFilter(ext));
        int i, n;
        n=list.length;
        String stri, stro;
        for(i=0; i<n; i++) {
            stri=nameSrc + list[i]; // имя файла в исходном каталоге
            stro=nameDst + list[i]; // имя файла в выходном каталоге
            File fileIn=new File(stri);
            File fileOut=new File(stro);
            if(fileIn.renameTo(fileOut)) {
                cnt++;
            }else {
                Log("?-ERROR-File don't move to: " + stro);
            }
        } // end for
        return cnt;
    }

}  // end TableDOCX(

