package spreadsheet.app;

import domain.Spreadsheet;
import java.io.FileNotFoundException;

public class Main {
		 public static void main(String[] args) throws FileNotFoundException {
        Spreadsheet spr = new Spreadsheet(10, 10);

        spr.put("A3", "123");
        spr.put("A2", "1");

        spr.put("B9", "=41+A2");
        spr.put("J5", "=7*6");
        spr.put("J6", "=3/2");            // Ganzzahl-Division: 1

        spr.put("C1", "=SUMME(A2:A3)");   // 124
        spr.put("C2", "=MAX(A2:A3)");     // 123
        spr.put("C3", "=MITTELWERT(A2:A3)"); // 62 (gerundet)
		spr.put("C5","=SUMME(C1:C3)");

        System.out.println(spr);
        spr.saveCsv("/tmp/test.csv");
    }
}
