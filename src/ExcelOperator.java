
import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelOperator {

    private static Integer[] rowNumbers = new Integer[]{11,13,15,17,18,20,21,23,24,26,27,28,29,30,32,33,34,36,38,39,40,41,42,43,44,46,49,50,51,55,
                                                 57,59,61,63,65,66,68};
    private static String directoryPath = "/Users/vmkri/Desktop/excel_files/";

    private static String[] projectNames = new String[]{"Nallavagu","Ghanapur Anicut(DS)","Ramadugu","Pocharam","Koulasanala","Sathanala",
                                                        "Mathadivagu","Swarna ","Suddavagu Gaddennavagu","Vattivagu","N.T.R.Sagar  ","P P Rao Project",
                                                        "Sri Komaram Bheem Project","Peddavagu Jagannathpur","Gollavagu","Nilwai Project","Ralivagu",
                                                        "Shanigaram","Boggulavagu","Malluruvagu","Lakhnavaram","Ramappa Lake","Palemvagu","Gundlavagu",
                                                        "Modikuntavagu","Upper manair","Peddavagu","Taliperu","Kinnerasani","Dindi","Asifnahar(DS)","Musi",
                                                        "Kotipally vagu","Pakhala Lake","Wyra","Lanka sagar","Bayyaram Tank"};

    private static String outputFile = "output.xlsx";

    private static Integer columnValue = 10;
    private static Map<String, List<Object[]>> projectWiseData = new HashMap<>();
    private static String filePrefix = "W_l._ "; //01.01.2017
    private static String fileSuffix = ".xlsx";
    private static String dateFormat = "dd.MM.yyyy";
    private static List<List<String>> listStrings = new ArrayList<>();

    private static void populateListStrings(String mon, String start, String end) {
        List<String> temp = new ArrayList<>();
        temp.add(mon);
        temp.add( start);
        temp.add(end);
        listStrings.add(temp);
    }
    public static void main(String args[]) throws IOException{
        if(args.length==1) {
            parseConfigFile(args[0]);
        }
        else {
            populateListStrings("Jan-2017", "2017-01-01", "2017-01-31");
            populateListStrings("Feb-2017", "2017-02-01", "2017-02-28");
            populateListStrings("Mar-2017", "2017-03-01", "2017-03-31");
            populateListStrings("Apr-2017", "2017-04-01", "2017-04-30");
            populateListStrings("May-2017", "2017-05-01", "2017-05-31");
        }

        generateOutputExcel();
        for(int i=0;i<listStrings.size();i++) {
            List<String> details = listStrings.get(i);
            getDataForDates(details.get(1), details.get(2));
            updateWorkbook(details.get(0), i+1);
            projectWiseData.clear();
        }
    }

    static void parseConfigFile(String configFile) throws FileNotFoundException {
        Gson gson  = new Gson();
        InputConfig input = gson.fromJson(new FileReader(configFile), InputConfig.class);
        rowNumbers = input.getProjectRows();
        directoryPath = input.getDirectoryPath();
        outputFile = input.getOutputFile();
        columnValue = input.getValueColumn();
        filePrefix = input.getPrefix();
        dateFormat = input.getDatePatternInFile();
        listStrings = input.getInputDates();
    }

    static void getDataForDates(String startDate, String endDate) throws IOException {


        LocalDate start = LocalDate.parse(startDate),
            end   = LocalDate.parse(endDate);

        LocalDate next = start.minusDays(1);
        while ((next = next.plusDays(1)).isBefore(end.plusDays(1))) {
            String date = new SimpleDateFormat(dateFormat).format(Date.from(next.atStartOfDay().atZone(ZoneId.systemDefault())
                                                                                  .toInstant()));
            Map<String, Double> details = getInfoFromFile(directoryPath+filePrefix+date+fileSuffix);
            for (Map.Entry<String, Double> entry : details.entrySet())
            {
                Object[] temp = new Object[]{date, entry.getValue()};
                if(projectWiseData.containsKey(entry.getKey())) {
                    List<Object[]> value = projectWiseData.get(entry.getKey());
                    value.add(temp);
                    projectWiseData.put(entry.getKey(), value);
                }
                else {
                    List<Object[]> value = new ArrayList<>();
                    value.add(temp);
                    projectWiseData.put(entry.getKey(), value);
                }
            }
        }
    }

    static void updateWorkbook(String rowValue, int rownumber) throws IOException {
        FileInputStream file = new FileInputStream(new File(directoryPath+outputFile));
        XSSFWorkbook workbook = new XSSFWorkbook (file);
        for (Map.Entry<String, List<Object[]>> entry : projectWiseData.entrySet()) {
            XSSFSheet sheet = workbook.getSheet(entry.getKey());
            if(sheet == null) {
                sheet = workbook.createSheet(entry.getKey());
            }
            Row row = sheet.getRow(0);
            if(row==null) {
                sheet.createRow(0);
                for(int i=1; i<=31; i++) {
                    Cell cell = sheet.getRow(0).createCell(i);
                    cell.setCellValue(i);
                }
            }
            sheet.createRow(rownumber).createCell(0).setCellValue(rowValue);
            List<Object[]> values = entry.getValue();
            for(int i=0;i<values.size(); i++) {
                Object[] objectArr = values.get(i);
                sheet.getRow(rownumber).createCell(Integer.parseInt(objectArr[0].toString().substring(0,2)))
                     .setCellValue((Double)objectArr[1]);
            }
        }
        file.close();

        FileOutputStream outFile =new FileOutputStream(new File(directoryPath+outputFile));
        workbook.write(outFile);
        outFile.close();
    }

    static void generateOutputExcel() throws IOException{
        FileOutputStream out =
            new FileOutputStream(new File(directoryPath+outputFile));
        XSSFWorkbook workbook = new XSSFWorkbook();
        workbook.write(out);
        out.close();
    }

    //static void writeInfoToFile(String fileName)

    static Map<String, Double> getInfoFromFile(String fileName) throws IOException {
        Map<String, Double> values = new HashMap<>();
        XSSFSheet sheet = getSheetData(fileName);

        for(int i=0;i<projectNames.length; i++) {
            int rowNum = findRow(sheet, projectNames[i]);
            if(rowNum!=-1) {
                Row row = sheet.getRow(rowNum);
                double capacity = row.getCell(columnValue).getCellTypeEnum() == CellType.STRING ? 0 : row.getCell(columnValue).getNumericCellValue();
                values.put(projectNames[i], capacity);
            }
        }
        /*for(int i=0;i<rowNumbers.length; i++) {
            Row row = sheet.getRow(rowNumbers[i]-1);
            double capacity = row.getCell(columnValue).getCellTypeEnum() == CellType.STRING ? 0 : row.getCell(columnValue).getNumericCellValue();
            values.put(row.getCell(1).getStringCellValue(), capacity);
            System.out.print("\""+row.getCell(1)+"\",");
        }*/
        return values;
    }

    static int findRow(XSSFSheet sheet, String cellContent) {
        for (Row row : sheet) {
            Cell cell = row.getCell(1);
            if (cell!=null && cell.getCellTypeEnum() == CellType.STRING) {
                if (cell.getRichStringCellValue().getString().trim().equals(cellContent.trim())) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    static XSSFSheet getSheetData(String fileName) throws IOException {
        FileInputStream file = new FileInputStream(new File(fileName));
        XSSFWorkbook workbook = new XSSFWorkbook (file);
        return workbook.getSheetAt(0);
    }
}
