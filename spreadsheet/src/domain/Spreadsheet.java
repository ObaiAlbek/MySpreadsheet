package domain;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Core logic class for the Spreadsheet application.
 * Handles cell storage, parsing, formula evaluation, and file I/O.
 */
public class Spreadsheet {

    private final Cell[][] cells;
    private final int rows;
    private final int cols;

    // Regex pattern to tokenize formulas into: Cell References, Ranges, Numbers, Operators, Parentheses, etc.
    private static final Pattern TOKEN_PATTERN = Pattern.compile(
            "([A-Z]+[0-9]+)|(:)|([0-9]+)|([+\\-*/^])|([()])|(,)"
    );

    /**
     * Initializes a new empty spreadsheet with specific dimensions.
     */
    public Spreadsheet(int rows, int cols) {
        this.rows = rows;
        this.cols = cols;

        cells = new Cell[rows][cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                cells[r][c] = new Cell();
    }

    /**
     * Gets the value of a cell by its address string (e.g., "A1").
     */
    public String get(String cellName) {
        int[] rc = parseAddress(cellName);
        return get(rc[0], rc[1]);
    }

    /**
     * Puts a value or formula into a specific cell by address (e.g., "A1").
     * Triggers evaluation if the input starts with "=".
     */
    public void put(String cellName, String value) {
        int[] rc = parseAddress(cellName);
        put(rc[0], rc[1], value);
    }

    /**
     * Imports data from a CSV file into the spreadsheet.
     * * @param path The file path to the CSV.
     * @param separator The delimiter (e.g., ',' or ';').
     * @param startCellName The top-left cell where insertion begins (e.g., "A1").
     */
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
                    
                    // Identify if it is a formula or a raw value
                    if (s.startsWith("=")) {
                        cells[r][c].setFormula(s);
                        evaluateCell(r, c); // Recalculate immediately
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

    /**
     * Exports the current state to a CSV file.
     * Saves formulas (starting with '=') if they exist, otherwise saves the calculated value.
     */
    public void saveCsv(String path) throws FileNotFoundException {
        try (PrintWriter out = new PrintWriter(path)) {
            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < cols; c++) {
                    Cell cell = cells[r][c];
                    // Save formula if present, otherwise value
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

    // --- Private Helper Methods ---

    private String get(int row, int col) {
        return cells[row][col].getValue();
    }

    private void put(int row, int col, String input) {
        input = input == null ? "" : input.trim();
        if (!input.startsWith("=")) {
            // It's a raw value (text or number)
            cells[row][col].setFormula("");
            cells[row][col].setValue(input);
        } else {
            // It's a formula
            cells[row][col].setFormula(input);
            evaluateCell(row, col);
        }
    }

    /**
     * Converts a cell address (e.g., "B2") into row and column indices.
     * "B2" -> Row 1, Col 1 (0-based).
     */
    private int[] parseAddress(String cellName) {
        if (cellName == null) throw new IllegalArgumentException("null address");
        String s = cellName.trim().toUpperCase();
        
        // Regex to separate Letters (Col) and Numbers (Row)
        Matcher m = Pattern.compile("^([A-Z]+)([0-9]+)$").matcher(s);
        if (!m.find()) throw new IllegalArgumentException("Invalid cell address: " + cellName);

        String colStr = m.group(1);
        String rowStr = m.group(2);

        int col = colStringToIndex(colStr);
        int row = Integer.parseInt(rowStr) - 1; // Convert 1-based to 0-based

        if (row < 0 || row >= rows || col < 0 || col >= cols)
            throw new IllegalArgumentException("Address out of bounds: " + cellName);

        return new int[]{row, col};
    }

    // Converts column letters (A, B... Z) to indices (0, 1... 25).
    // MVP limitation: Only supports single letters.
    private int colStringToIndex(String colStr) {
        if (colStr.length() != 1) throw new IllegalArgumentException("Only A..Z supported");
        char ch = colStr.charAt(0);
        if (ch < 'A' || ch > 'Z') throw new IllegalArgumentException("Invalid column: " + colStr);
        return ch - 'A';
    }

    // Parses a range string (e.g., "A1:A5") into coordinates {r1, c1, r2, c2}.
    private int[] parseRange(String range) {
        String s = range.trim().toUpperCase();
        Matcher m = Pattern.compile("^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$").matcher(s);
        if (!m.find()) throw new IllegalArgumentException("Invalid range: " + range);
        
        int[] a = parseAddress(m.group(1));
        int[] b = parseAddress(m.group(2));
        
        // Ensure standard order (top-left to bottom-right)
        int r1 = Math.min(a[0], b[0]), c1 = Math.min(a[1], b[1]);
        int r2 = Math.max(a[0], b[0]), c2 = Math.max(a[1], b[1]);
        return new int[]{r1, c1, r2, c2};
    }

    private String[] splitCsv(String line, char sep) {
        return line.split(Pattern.quote(String.valueOf(sep)), -1);
    }

    /**
     * The core logic engine. Determines if the formula is a function (SUM, MIN)
     * or a mathematical expression, evaluates it, and handles errors.
     */
    private void evaluateCell(int row, int col) {
        String f = cells[row][col].getFormula().trim();
        String result;

        try {
            // Check for predefined functions
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
                // Not a function, try to evaluate as a math expression (e.g., A1+5)
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

    // Extracts the content inside parentheses: SUMME(A1:A2) -> A1:A2
    private String insideOf(String f, String funcName) {
        if (!f.endsWith(")")) throw new IllegalArgumentException("Invalid function syntax: " + f);
        int start = funcName.length() + 1;
        return f.substring(start, f.length() - 1).trim();
    }

    // --- Aggregation Functions ---

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

    // Helper to get all numeric values from a range or a single cell reference.
    private Iterable<Long> valuesOf(String rangeOrList) {
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

    // --- Expression Evaluation (Shunting-yard Algorithm) ---

    private String evalExpression(String expr) {
        List<String> rpn = toRPN(expr); // Convert Infix to Postfix
        long val = evalRPN(rpn);        // Calculate Postfix
        return String.valueOf(val);
    }

    /**
     * Converts an infix expression (3 + 4) to Reverse Polish Notation (3 4 +)
     * using the Shunting-yard algorithm.
     */
    private List<String> toRPN(String expr) {
        expr = expr.toUpperCase().replaceAll("\\s+", "");
        Matcher m = TOKEN_PATTERN.matcher(expr);

        List<String> output = new ArrayList<>();
        Deque<String> ops = new ArrayDeque<>();

        while (m.find()) {
            String t;
            if ((t = m.group(1)) != null) {         // Cell Reference (e.g., A1)
                output.add(resolveRef(t));
            } else if ((t = m.group(3)) != null) {  // Number (e.g., 100)
                output.add(t);
            } else if ((t = m.group(4)) != null) {  // Operator (+, -, *, /)
                while (!ops.isEmpty() && isOperator(ops.peek()) &&
                        (precedence(ops.peek()) > precedence(t) ||
                         (precedence(ops.peek()) == precedence(t) && isLeftAssoc(t)))) {
                    output.add(ops.pop());
                }
                ops.push(t);
            } else if ((t = m.group(5)) != null) {  // Parentheses
                if (t.equals("(")) ops.push(t);
                else {
                    while (!ops.isEmpty() && !ops.peek().equals("(")) output.add(ops.pop());
                    if (ops.isEmpty() || !ops.peek().equals("(")) throw new IllegalArgumentException("Mismatched parens");
                    ops.pop();
                }
            } else if (m.group(2) != null || m.group(6) != null) {
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

    /**
     * Evaluates a Reverse Polish Notation (RPN) list.
     * Uses a stack to process operands and operators.
     */
    private long evalRPN(List<String> rpn) {
        Deque<Long> st = new ArrayDeque<>();
        for (String t : rpn) {
            if (isOperator(t)) {
                if (st.size() < 2) throw new IllegalArgumentException("Malformed expression");
                long b = st.pop(), a = st.pop();
                switch (t) {
                    case "+" -> st.push(a + b);
                    case "-" -> st.push(a - b);
                    case "*" -> st.push(a * b);
                    case "/" -> {
                        if (b == 0) throw new ArithmeticException("/0");
                        st.push(a / b);
                    }
                    case "^" -> st.push((long)Math.pow(a, b));
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
    
    // Defines Operator Precedence (BODMAS / PEMDAS)
    private int precedence(String op) {
        return switch (op) {
            case "^" -> 3; // Highest
            case "*", "/" -> 2;
            case "+", "-" -> 1; // Lowest
            default -> 0;
        };
    }
    
    private boolean isLeftAssoc(String op) {
        return !op.equals("^"); // Exponentiation is typically right-associative
    }

    /**
     * Resolves a cell reference (e.g., "A1") to its numeric value.
     * Used during expression evaluation.
     */
    private String resolveRef(String ref) {
        int[] rc = parseAddress(ref);
        String v = cells[rc[0]][rc[1]].getValue().trim();
        if (v.isEmpty()) return "0";
        if (v.startsWith("#")) throw new IllegalArgumentException("Ref error: " + ref);
        parseLongStrict(v); // Validate it's a number
        return v;
    }

    private long parseLongStrict(String s) {
        if (!s.matches("[-]?[0-9]+")) throw new NumberFormatException("Not an integer: " + s);
        return Long.parseLong(s);
    }

    /**
     * Returns a string representation of the entire grid for console display.
     * Includes row numbers and column headers.
     */
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("    ");
        // Print Column Headers (A, B, C...)
        for (int i = 0; i < cols; i++) sb.append("  ").append((char)('A' + i)).append("  | ");
        int rc = 1;
        // Print Rows
        for (int r = 0; r < rows; r++) {
            sb.append(System.lineSeparator());
            sb.append(String.format("%2s", rc++)).append(": "); // Row Number
            for (int c = 0; c < cols; c++) sb.append(cells[r][c]).append(" | ");
        }
        return sb.toString();
    }
}