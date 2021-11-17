package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

    public class excelmain {
        //constructed

        public static void main(String[] args) throws IOException {
            getCellData();
            getRowCount();
              }

        public static void getCellData() throws IOException {
            String excelPath = "./data/TestData.xlsx";
            XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
            XSSFSheet sheet = workbook.getSheet("sheet1");

            DataFormatter formatter = new DataFormatter();
            Object value = formatter.formatCellValue(sheet.getRow(1).getCell(2));

            // double value = sheet.getRowBreaks(1).getClass(0).getNumericCellValue();d
            System.out.println(value);
        }
        public static void getRowCount() throws IOException {

          try{
            String excelPath = "./data/TestData.xlsx";
            XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
            XSSFSheet sheet = workbook.getSheet("sheet1");

            int rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println("No of Rows :" + rowCount);

} catch (Exception exp) {
            System.out.println(exp.getCause());
            System.out.println(exp.getMessage());
            exp.printStackTrace();
             }
        }
                }


