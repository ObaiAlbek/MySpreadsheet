package domain;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Spreadsheet {

    private final Cell[][] cells;
    private final int rows;
    private final int cols;

    // Tokenizer: Zellref, Bereichs-Doppelpunkt, Zahl, Operator, Klammer, Komma
    private static final Pattern TOKEN_PATTERN = Pattern.compile(
            "([A-Z]+[0-9]+)|(:)|([0-9]+)|([+\\-*/^])|([()])|(,)"
    );

    public Spreadsheet(int rows, int cols) {
        if (rows < 1 || rows > 99) throw new IllegalArgumentException("rows must be 1..99");
        if (cols < 1 || cols > 26) throw new IllegalArgumentException("cols must be 1..26 (A..Z)");
        this.rows = rows;
        this.cols = cols;

        cells = new Cell[rows][cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                cells[r][c] = new Cell();
    }

    // ---------------------------
    // Public API

    public String get(String cellName) {
        int[] rc = parseAddress(cellName);
        return get(rc[0], rc[1]);
    }

    public void put(String cellName, String value) {
        int[] rc = parseAddress(cellName);
        put(rc[0], rc[1], value);
    }

    public void readCsv(String path, char separator, String startCellName) throws FileNotFoundException {
        int[] rc = parseAddress(startCellName);
        int startR = rc[0], startC = rc[1];

        try (BufferedReader br = new BufferedReader(new FileReader(path))) {
            String line;
            int r = startR;
            while ((line = br.readLine()) != null && r < rows) {
                String[] parts = splitCsv(line, separator);
                int c = startC;
                for (String raw : parts) {
                    if (c >= cols) break;
                    String s = raw.trim();
                    if (s.startsWith("=")) {
                        cells[r][c].setFormula(s);
                        evaluateCell(r, c);
                    } else {
                        cells[r][c].setFormula("");
                        cells[r][c].setValue(s);
                    }
                    c++;
                }
                r++;
            }
        } catch (IOException e) {
            throw new FileNotFoundException(e.getMessage());
        }
    }

    public void saveCsv(String path) throws FileNotFoundException {
        try (PrintWriter out = new PrintWriter(path)) {
            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < cols; c++) {
                    Cell cell = cells[r][c];
                    if (!cell.getFormula().isEmpty())
                        out.print("=" + cell.getFormula());
                    else
                        out.print(cell.getValue());
                    if (c < cols - 1) out.print(",");
                }
                out.println();
            }
        }
    }

    // ---------------------------
    // Internals

    private String get(int row, int col) {
        return cells[row][col].getValue();
    }

    private void put(int row, int col, String input) {
        input = input == null ? "" : input.trim();
        if (!input.startsWith("=")) {
            cells[row][col].setFormula("");
            cells[row][col].setValue(input);
        } else {
            cells[row][col].setFormula(input);
            evaluateCell(row, col);
        }
    }

    // A1-Adressparser → [rowIndex, colIndex]
    private int[] parseAddress(String cellName) {
        if (cellName == null) throw new IllegalArgumentException("null address");
        String s = cellName.trim().toUpperCase();
        Matcher m = Pattern.compile("^([A-Z]+)([0-9]+)$").matcher(s);
        if (!m.find()) throw new IllegalArgumentException("Invalid cell address: " + cellName);

        String colStr = m.group(1);
        String rowStr = m.group(2);

        int col = colStringToIndex(colStr);
        int row = Integer.parseInt(rowStr) - 1;

        if (row < 0 || row >= rows || col < 0 || col >= cols)
            throw new IllegalArgumentException("Address out of bounds: " + cellName);

        return new int[]{row, col};
    }

    // Nur A..Z (1..26) für dieses MVP
    private int colStringToIndex(String colStr) {
        if (colStr.length() != 1) throw new IllegalArgumentException("Only A..Z supported");
        char ch = colStr.charAt(0);
        if (ch < 'A' || ch > 'Z') throw new IllegalArgumentException("Invalid column: " + colStr);
        return ch - 'A';
    }

    // Bereichsparser "A1:B3" → [r1,c1,r2,c2] (geordnet)
    private int[] parseRange(String range) {
        String s = range.trim().toUpperCase();
        Matcher m = Pattern.compile("^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$").matcher(s);
        if (!m.find()) throw new IllegalArgumentException("Invalid range: " + range);
        int[] a = parseAddress(m.group(1));
        int[] b = parseAddress(m.group(2));
        int r1 = Math.min(a[0], b[0]), c1 = Math.min(a[1], b[1]);
        int r2 = Math.max(a[0], b[0]), c2 = Math.max(a[1], b[1]);
        return new int[]{r1, c1, r2, c2};
    }

    // CSV splitter (einfach, ohne Quotes)
    private String[] splitCsv(String line, char sep) {
        return line.split(Pattern.quote(String.valueOf(sep)), -1);
    }

    // ---------------------------
    // Evaluation

    private void evaluateCell(int row, int col) throws NumberFormatException {
        String f = cells[row][col].getFormula().trim();
        String result = "";

        try {
            if (f.toUpperCase().startsWith("SUMME(")) {
                String inner = insideOf(f, "SUMME");
                result = String.valueOf(sum(inner));
            } else if (f.toUpperCase().startsWith("MIN(")) {
                String inner = insideOf(f, "MIN");
                result = String.valueOf(min(inner));
            } else if (f.toUpperCase().startsWith("MAX(")) {
                String inner = insideOf(f, "MAX");
                result = String.valueOf(max(inner));
            } else if (f.toUpperCase().startsWith("MITTELWERT(")) {
                String inner = insideOf(f, "MITTELWERT");
                result = String.valueOf(avg(inner));
            } else if (!f.isEmpty()) {
                result = evalExpression(f);
            } else {
                result = "";
            }
        } catch (ArithmeticException ae) {
            result = "#DIV/0!";
        } catch (IllegalArgumentException ex) {
            result = "#ERR";
        }

        cells[row][col].setValue(result);
    }

    private String insideOf(String f, String funcName) {
        if (!f.endsWith(")")) throw new IllegalArgumentException("Invalid function syntax: " + f);
        int start = funcName.length() + 1;
        return f.substring(start, f.length() - 1).trim();
    }

    // ---- Bereichs-Funktionen (ein Argument: Range)
    private long sum(String arg) {
        long s = 0;
        for (long v : valuesOf(arg)) s += v;
        return s;
    }

    private long min(String arg) {
        long best = Long.MAX_VALUE;
        boolean any = false;
        for (long v : valuesOf(arg)) { best = Math.min(best, v); any = true; }
        if (!any) throw new IllegalArgumentException("Empty range");
        return best;
    }

    private long max(String arg) {
        long best = Long.MIN_VALUE;
        boolean any = false;
        for (long v : valuesOf(arg)) { best = Math.max(best, v); any = true; }
        if (!any) throw new IllegalArgumentException("Empty range");
        return best;
    }

    private long avg(String arg) {
        long s = 0; long n = 0;
        for (long v : valuesOf(arg)) { s += v; n++; }
        if (n == 0) throw new IllegalArgumentException("Empty range");
        return Math.round((double)s / (double)n);
    }

    private Iterable<Long> valuesOf(String rangeOrList) {
        // Erlaubt aktuell nur EINEN Range wie "A1:B3"
        int[] r = parseRange(rangeOrList);
        List<Long> vals = new ArrayList<>();
        for (int rr = r[0]; rr <= r[2]; rr++) {
            for (int cc = r[1]; cc <= r[3]; cc++) {
                String txt = cells[rr][cc].getValue().trim();
                if (txt.isEmpty()) continue;
                vals.add(parseLongStrict(txt));
            }
        }
        return vals;
    }

    // ---- Ausdrucksauswertung (Shunting-Yard → RPN → Eval)
    private String evalExpression(String expr) {
        List<String> rpn = toRPN(expr);
        long val = evalRPN(rpn);
        return String.valueOf(val);
    }

    private List<String> toRPN(String expr) {
        expr = expr.toUpperCase().replaceAll("\\s+", "");
        Matcher m = TOKEN_PATTERN.matcher(expr);

        List<String> output = new ArrayList<>();
        Deque<String> ops = new ArrayDeque<>();

        while (m.find()) {
            String t;
            if ((t = m.group(1)) != null) {               // Zellreferenz
                output.add(resolveRef(t));
            } else if ((t = m.group(3)) != null) {        // Zahl
                output.add(t);
            } else if ((t = m.group(4)) != null) {        // Operator
                while (!ops.isEmpty() && isOperator(ops.peek()) &&
                        (precedence(ops.peek()) > precedence(t) ||
                         (precedence(ops.peek()) == precedence(t) && isLeftAssoc(t)))) {
                    output.add(ops.pop());
                }
                ops.push(t);
            } else if ((t = m.group(5)) != null) {        // Klammer
                if (t.equals("(")) ops.push(t);
                else {
                    while (!ops.isEmpty() && !ops.peek().equals("(")) output.add(ops.pop());
                    if (ops.isEmpty() || !ops.peek().equals("(")) throw new IllegalArgumentException("Mismatched parens");
                    ops.pop(); // '('
                }
            } else if (m.group(2) != null || m.group(6) != null) {
                // ':' und ',' sind in reinen Ausdrücken nicht erlaubt
                throw new IllegalArgumentException("Unexpected token: ':' or ',' in expression");
            }
        }
        while (!ops.isEmpty()) {
            String op = ops.pop();
            if (op.equals("(")) throw new IllegalArgumentException("Mismatched parens");
            output.add(op);
        }
        return output;
    }

    private long evalRPN(List<String> rpn) {
        Deque<Long> st = new ArrayDeque<>();
        for (String t : rpn) {
            if (isOperator(t)) {
                if (st.size() < 2) throw new IllegalArgumentException("Malformed expression");
                long b = st.pop(), a = st.pop();
                switch (t) {
                    case "+": st.push(a + b); break;
                    case "-": st.push(a - b); break;
                    case "*": st.push(a * b); break;
                    case "/":
                        if (b == 0) throw new ArithmeticException("/0");
                        st.push(a / b); break;
                    case "^":
                        st.push((long)Math.pow(a, b)); break;
                }
            } else {
                st.push(parseLongStrict(t));
            }
        }
        if (st.size() != 1) throw new IllegalArgumentException("Malformed expression");
        return st.pop();
    }

    private boolean isOperator(String s) {
        return "+-*/^".contains(s);
    }
    private int precedence(String op) {
        switch (op) {
            case "^": return 3;
            case "*": case "/": return 2;
            case "+": case "-": return 1;
            default: return 0;
        }
    }
    private boolean isLeftAssoc(String op) {
        return !op.equals("^"); // Potenz ist rechtsassoziativ
    }

    private String resolveRef(String ref) {
        int[] rc = parseAddress(ref);
        String v = cells[rc[0]][rc[1]].getValue().trim();
        if (v.isEmpty()) return "0";
        // Falls in der referenzierten Zelle ein Fehlertext steht, brich ab:
        if (v.startsWith("#")) throw new IllegalArgumentException("Ref error: " + ref);
        // Nur Ganzzahlen erlaubt (MVP)
        parseLongStrict(v); // Validierung
        return v;
    }

    private long parseLongStrict(String s) {
        if (!s.matches("[-]?[0-9]+")) throw new NumberFormatException("Not an integer: " + s);
        return Long.parseLong(s);
    }

    // ---------------------------

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("    ");
        for (int i = 0; i < cols; i++) sb.append("  ").append((char)('A' + i)).append("  | ");
        int rc = 1;
        for (int r = 0; r < rows; r++) {
            sb.append(System.lineSeparator());
            sb.append(String.format("%2s", rc++)).append(": ");
            for (int c = 0; c < cols; c++) sb.append(cells[r][c]).append(" | ");
        }
        return sb.toString();
    }
}
