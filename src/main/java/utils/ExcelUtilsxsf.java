package utils;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ExcelUtilsxsf {

        static HSSFWorkbook workbook;
        static HSSFSheet sheet;

        public ExcelUtilsxsf(String excelPath, String sheetName) throws IOException
        {
            InputStream file =new FileInputStream(excelPath)
            workbook = new HSSFWorkbook(new POIFSFileSystem(file));
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


