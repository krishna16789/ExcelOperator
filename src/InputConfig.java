import java.util.List;

public class InputConfig {
    Integer[] projectRows;
    String directoryPath;
    String outputFile;
    List<List<String>> inputDates;
    String prefix;
    String datePatternInFile;
    Integer valueColumn;

    public Integer getValueColumn() {
        return valueColumn;
    }

    public void setValueColumn(Integer valueColumn) {
        this.valueColumn = valueColumn;
    }

    public Integer[] getProjectRows() {
        return projectRows;
    }

    public void setProjectRows(Integer[] projectRows) {
        this.projectRows = projectRows;
    }

    public String getDirectoryPath() {
        return directoryPath;
    }

    public void setDirectoryPath(String directoryPath) {
        this.directoryPath = directoryPath;
    }

    public String getOutputFile() {
        return outputFile;
    }

    public void setOutputFile(String outputFile) {
        this.outputFile = outputFile;
    }

    public List<List<String>> getInputDates() {
        return inputDates;
    }

    public void setInputDates(List<List<String>> inputDates) {
        this.inputDates = inputDates;
    }

    public String getPrefix() {
        return prefix;
    }

    public void setPrefix(String prefix) {
        this.prefix = prefix;
    }

    public String getDatePatternInFile() {
        return datePatternInFile;
    }

    public void setDatePatternInFile(String datePatternInFile) {
        this.datePatternInFile = datePatternInFile;
    }
}
