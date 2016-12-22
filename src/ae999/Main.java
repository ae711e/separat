package ae999;

import java.io.File;

public class Main {

    public static void main(String[] args) {
	// write your code here
        if (args.length<1) {
            System.out.println("(C) 2016 Àëåêñåé Åðåìèí");
            System.out.println("ÐÀÇÄÅËÈÌ ÊÀÐÒÈÍÊÈ");
            System.out.println("SEPARAT docfile.RTF");
            return ;
        }

        TableDOCX doc=new TableRTF();
        boolean b;
        File f = new File(args[0]);
        if(f.exists()) {
            b = doc.parse(f);
        } else {
            System.out.println("?-Error-File not found: " + f.getAbsolutePath());
        }
    }
}
