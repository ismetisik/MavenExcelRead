package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ExcelUtils {
    //constructed

    static XSSFWorkbook workbook;
    static XSSFSheet sheet;

    public ExcelUtils(String excelPath, String sheetName) throws IOException {
            workbook = new XSSFWorkbook(excelPath);
            sheet = workbook.getSheet("sheet1");
    }

    public static void getCellData(int rowNum,int colNum ) throws IOException {

/* we do not need following items since we have constructed variables in the class level
        String excelPath = "./data/TestData.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
        XSSFSheet sheet = workbook.getSheet("sheet1");
*/
        DataFormatter formatter = new DataFormatter();
        Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
        System.out.println(value);
    }

    public static void getRowCount() {

            int rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println("No of Rows :" + rowCount);


        //try} catch (Exception exp) {
          //  System.out.println(exp.getCause());
            //System.out.println(exp.getMessage());
            //exp.printStackTrace();
       // }
    }
}
