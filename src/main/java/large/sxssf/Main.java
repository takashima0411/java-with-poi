package large.sxssf;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;


public class Main {

    public static void main(String[] args) {
        Workbook wb = new SXSSFWorkbook();
        Sheet sheet1 = wb.createSheet("sheet 1");

        for (int i = 0; i < 10000; i++) {
            Row r = sheet1.createRow(i);
            for (int j = 0; j < 10000; j++) {
                r.createCell(j).setCellValue(i * j);
            }
        }

        try (FileOutputStream fos = new FileOutputStream("./workbook.xlsx")) {
            wb.write(fos);
        } catch (IOException ioe) {
            System.out.println(ioe.getMessage());
        }
    }

}
