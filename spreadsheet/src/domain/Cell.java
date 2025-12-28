package domain;

/**
 * Represents a single cell in the spreadsheet.
 * It stores both the raw formula (if applicable) and the evaluated display value.
 */
public class Cell {
    // The formula string (without the leading '='), e.g., "A1+B2"
    private String formula = "";
    
    // The current evaluated value of the cell, e.g., "42"
    private String value = "";
    
    public String getFormula() {
        return formula;
    }
    
    /**
     * Sets the formula for the cell.
     * automatically removes the leading '=' and normalizes to uppercase.
     * @param formula The raw input string (e.g. "=a1+b2")
     */
    public void setFormula(String formula) {
        if (!formula.isEmpty())
            // Remove leading '=' (index 0) and normalize case
            this.formula = formula.toUpperCase().substring(1);  
    }
    
    public String getValue() {
        return value;
    }
    
    public void setValue(String value) {
        this.value = value;
    }
    
    /**
     * Returns a formatted string representation of the value,
     * padded to 4 characters for console alignment.
     */
    public String toString() {
        return String.format("%4s", value);
    }
    
    public boolean isEmpty() {
        return value.isEmpty();
    }
}