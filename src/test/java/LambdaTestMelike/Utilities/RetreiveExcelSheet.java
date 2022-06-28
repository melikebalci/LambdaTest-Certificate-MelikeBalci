package LambdaTestMelike.Utilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class RetreiveExcelSheet {

    public static ArrayList<String> getData(String browserName) throws IOException {

        ArrayList<String> testCaseData = new ArrayList<>();

        FileInputStream fis = new FileInputStream(".//test_data//LambdaTestData.xlsx");

        XSSFWorkbook excel = new XSSFWorkbook(fis);

        int sheets = excel.getNumberOfSheets();

        for (int i = 0; i < sheets; i++) {
            if (excel.getSheetAt(i).getSheetName().equalsIgnoreCase("Sheet1")) {


                XSSFSheet sheet = excel.getSheetAt(i);

                Iterator<Row> rows = sheet.iterator();

                Row firstRow = rows.next();

                Iterator<Cell> cells = firstRow.cellIterator();

                int columnIndex = 0;

                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    if (cell.getStringCellValue().equalsIgnoreCase("Browsers"))
                        break;

                    columnIndex++;
                }


                while (rows.hasNext()) {
                    Row r = rows.next();
                    Cell c1 = r.getCell(columnIndex);

                    if (c1!=null && c1.getStringCellValue().equalsIgnoreCase(browserName)) {
                        Iterator<Cell> cv = r.cellIterator();
                        while (cv.hasNext()) {
                            Cell c2 = cv.next();

                            if (c2.getCellTypeEnum() == CellType.NUMERIC)
                                testCaseData.add(Double.toString(c2.getNumericCellValue()));

                            else if (c2.getCellTypeEnum() == CellType.STRING)
                                testCaseData.add(c2.getStringCellValue());
                        }

                        break;
                    }
                }
            }
        }

        excel.close();
        return testCaseData;
    }
}
