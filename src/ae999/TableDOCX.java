package ae999;

/*
 * Created by ������� on 03.07.2016.
 * (C) 2016, Aleksey Eremin
 *
 * ��������� ������� � �����  .docx
*/
// Modify:
// 2016-07-03 ������� ������� "�������� ������" emptedStarngeCell()
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

// ���������� ������� �� ����������� � ����� DOCX
public class TableDOCX {

    final protected static String s_separatorDir = System.getProperty("file.separator");
    final protected static String s_subdirJpg1 ="yYy";
    final protected static String s_subdirJpg2 ="x";
    
    boolean parse(File iFile)
    {
        final String strNEED_TEXT="������������ ����� ���������.";   	// �������� ��������� �������
        String str;
        int cn, cnt;
        int i, n;
        boolean result=false;
        String DiRi;		// �������� �������� � ����������
        FileInputStream fileInputStream;
        XWPFDocument docx;
        try {
            fileInputStream = new FileInputStream(iFile);
            // ��������� ���� � ��������� ��� ���������� � ������ XWPFDocument
            docx = new XWPFDocument(OPCPackage.open(fileInputStream));  // word 2007
            // ������� ��� �������� �������� �����
            DiRi=iFile.getParent();  // �������� �������� � ����������
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
        XWPFTable mytable = null;  // ��������� ������� �� ������� ����������
        int mycol=0; // ������� � ������� ������
        // ��������� �� �������� ��������� � ������� ������ �������,
        // �� ���� �����, � ������� � ��������� ������� ��������� �������� needText
        tableIter = docx.getTablesIterator();
        while(tableIter.hasNext()) {
            table = tableIter.next();		// ��������� �������
            //printTable(table);
            cellList=table.getRow(0).getTableCells(); // ����� ����� ������ ������
            // ���� ������� ����������
            cn=cellList.size()-1;  // ������ ��������� ������ (�������)
            // ���� ������� �� �����������
            if (cn == 4) {
                str=cellList.get(cn).getText();  // ��������� ������� (��� ������ ���� ����� ������)
                if( str.regionMatches(0, strNEED_TEXT, 0, strNEED_TEXT.length()) ) { // ������� � ������� ������� �������� �������
                    mytable=table;	// �������� �������
                    mycol=cn;	// �������� ������ �������
                    break;
                }
            }
        } // end while
        //
        // ����� � ��������  - ����� �� �������?
        // �������, ���� ��������� ������� ������ 0, ������ �����
        cnt = 0;
        if (mycol>1) {
            // �������� ����������, ��� ������, ������� ���� � ���������
            String jpgDir=DiRi+s_separatorDir+s_subdirJpg1;
            File wd1=new File(jpgDir);
            wd1.mkdir();
            
            //
            ////printTable(mytable);
            // ������ �� ��������� ������� � ����� � ����� ������� ������ � ������� ��� ��������
            rowList = mytable.getRows();        // ������ ����� �������
            n = rowList.size();                                // ���-�� ����� �������
            // ������ ��������� ������ ������ � ����� �������,
            // � ���� ����� ��� - ������� ������
            for (i = n - 1; i > 0; i--) {
                row = rowList.get(i);                                        // ����� ������ � �������� i
                cellList = row.getTableCells();                // �� ������ �������� ������ ����� (�������)
                str = cellList.get(mycol).getText();        // ��������� ������� - ��� ����� � ���������
                // ���� ��������� (����������) ����� ����� .jpg
                if (str.endsWith(".jpg")) {  // str.regionMatches(l-4, ".jpg",0,4) l=str.length(); // ����� ������
                    // �������� ��� �����
                    iFile = new File(DiRi, str);
                    // �������� ���� �� ����� ����?
                    if (iFile.exists()) {
                        // ����� ���
                        if (MoveIt(iFile, s_subdirJpg1)) {                // ���������� ����
                            cnt++;
                        }
                        ////Log("Not exist " + cnt + ") " + str);
                    }  // if (!file.exists())
                } // end if jpg
                //
            }
            //
            // ���������� ���������� ����� �� ��������� �������
            // �������� ����������, ��� ������, ������� �� ������ � ��������
            String outDir=DiRi+s_separatorDir+s_subdirJpg2;
            File wd2=new File(outDir);
            wd2.mkdir();
            int a;
            String maskExt=".*jpg$"; // ���� ������ jpg �����
            MoveFiles(DiRi, outDir, maskExt);
            // ���������� �� ������� �������� � ��������
            MoveFiles(jpgDir, DiRi, maskExt);
            // ������ ���� �������
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

    // ����������� �������
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

    // ���� ���� - ���������� TRUE
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

    // ����������� �����  �� ����� ���������� (dir_src) � ������ (dir_dst) � �������� ����������� (ext)
    static int MoveFiles(String dir_src, String dir_dst, String ext)
    {
      int cnt=0;
        // �������� ���������� - ���� ��� � �������� �� �����������?
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
    
        // ������� ��� ����������, ��� ����� �����
        String nameSrc;
        nameSrc=dirSrc.getAbsolutePath() + s_separatorDir;  // ������� �����������
        // ������� ��� ����������, ���� ����� ������
        String nameDst;
        nameDst=dirDst.getAbsolutePath() + s_separatorDir;  // ������� �����������
        
        // �������� ��� Regexp ���������,
        String[] list=dirSrc.list(new MyFilter(ext));
        int i, n;
        n=list.length;
        String stri, stro;
        for(i=0; i<n; i++) {
            stri=nameSrc + list[i]; // ��� ����� � �������� ��������
            stro=nameDst + list[i]; // ��� ����� � �������� ��������
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

