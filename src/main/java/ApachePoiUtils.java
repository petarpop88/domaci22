import java.io.*;

import com.github.javafaker.Faker;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePoiUtils {

    public static void readAndPrint(String fileName){
        try{
            FileInputStream fileInputStream = new FileInputStream(new File(fileName));

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            for(int i = 0; i < 2; i++){
                XSSFRow row = sheet.getRow(i);

                for(int j = 0; j < 2;  j++){
                    XSSFCell cell = row.getCell(j);
                    String text = cell.getStringCellValue();
                    System.out.println(text + " ");
                }
            }
            System.out.println("-------------");
        }
        catch (IOException ex){
            System.out.println("FileNotFound.class");
        }
    }

    public static void write(String fileName){
        try{
            FileInputStream fileInputStream = new FileInputStream(new File(fileName));

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            Faker faker = new Faker();

            for(int i = 2; i < 10; i++){

                XSSFRow row = sheet.createRow(i);

                for(int j = 0; j < 2; j++){
                    XSSFCell cell1 = row.createCell(0);
                    cell1.setCellValue(faker.name().firstName());

                    XSSFCell cell2 = row.createCell(1);
                    cell2.setCellValue(faker.name().lastName());
                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
            workbook.write(fileOutputStream);
            fileOutputStream.close();

        }
        catch (IOException ex){
            System.out.println("FileNotFound.class");
        }
    }

    public static void readAllAndPrint(String fileName){
        try{
            FileInputStream fileInputStream = new FileInputStream(new File(fileName));

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            for(int i = 0; i < 10; i++){
                XSSFRow row = sheet.getRow(i);

                for(int j = 0; j < 2;  j++){
                    XSSFCell cell = row.getCell(j);
                    String text = cell.getStringCellValue();
                    System.out.println(text + " ");
                }
            }
            System.out.println("-------------");
        }
        catch (IOException ex){
            System.out.println("FileNotFound.class");
        }
    }
}
