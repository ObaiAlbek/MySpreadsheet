package domain;

public class Cell {
    private String formula = "";
    private String value = "";
    
    public String getFormula() {
        return formula;
    }
    
    public void setFormula(String formula) {
        if (!formula.isEmpty())
            this.formula = formula.toUpperCase().substring(1);  
    }
    
    public String getValue() {
        return value;
    }
    
    public void setValue(String value) {
        this.value = value;
    }
    
    public String toString() {
        return String.format("%4s", value);
    }
    
    public boolean isEmpty() {
        return value.isEmpty();
    }
    
}

