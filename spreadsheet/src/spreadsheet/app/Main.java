package spreadsheet.app;

import domain.Spreadsheet;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) {
        // Tabelle erstellen (10 Zeilen, 10 Spalten)
        Spreadsheet spr = new Spreadsheet(10, 10);
        Scanner scanner = new Scanner(System.in);
        
        System.out.println("=== Java Spreadsheet v1.0 ===");
        System.out.println("Befehle:");
        System.out.println("  put <Zelle> <Wert>   (z.B. put A1 100 oder put C1 =SUMME(A1:A2))");
        System.out.println("  get <Zelle>          (z.B. get A1)");
        System.out.println("  print                (Zeigt die ganze Tabelle)");
        System.out.println("  save <Datei>         (z.B. save test.csv)");
        System.out.println("  load <Datei>         (z.B. load test.csv)");
        System.out.println("  exit                 (Beenden)");
        System.out.println("-----------------------------");

        // 1. Automatisch drucken beim Start
        System.out.println(spr);

        while (true) {
            System.out.print("Eingabe> ");
            String input = scanner.nextLine().trim();

            // Leere Eingabe ignorieren
            if (input.isEmpty()) continue;

            // Eingabe zerlegen (max 3 Teile: Befehl, Arg1, Arg2)
            String[] parts = input.split("\\s+", 3);
            String command = parts[0].toLowerCase();

            try {
                switch (command) {
                    case "exit":
                        System.out.println("Programm beendet.");
                        scanner.close();
                        return; // Main verlassen

                    case "print":
                        // Nichts tun, Tabelle wird am Ende der Schleife eh gedruckt
                        break;

                    case "put":
                        if (parts.length < 3) {
                            System.out.println("Fehler: Bitte Zelle und Wert angeben (z.B. put A1 50)");
                        } else {
                            String cell = parts[1];
                            String value = parts[2];
                            spr.put(cell, value);
                            System.out.println("OK: " + cell + " = " + value);
                        }
                        break;

                    case "get":
                        if (parts.length < 2) {
                            System.out.println("Fehler: Bitte Zelle angeben (z.B. get A1)");
                        } else {
                            System.out.println("Wert: " + spr.get(parts[1]));
                        }
                        break;

                    case "save":
                        if (parts.length < 2) {
                            System.out.println("Fehler: Bitte Dateiname angeben (z.B. save daten.csv)");
                        } else {
                            spr.saveCsv(parts[1]);
                            System.out.println("Tabelle gespeichert in: " + parts[1]);
                        }
                        break;

                    case "load":
                        if (parts.length < 2) {
                            System.out.println("Fehler: Bitte Dateiname angeben (z.B. load daten.csv)");
                        } else {
                            spr.readCsv(parts[1], ',', "A1");
                            System.out.println("Tabelle geladen aus: " + parts[1]);
                        }
                        break;

                    default:
                        System.out.println("Unbekannter Befehl. Verfügbar: put, get, print, save, load, exit");
                }
            } catch (Exception e) {
                System.err.println("Fehler aufgetreten: " + e.getMessage());
            }

            System.out.println();
            System.out.println(spr);
        }
    }
}