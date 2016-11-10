package image;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import static org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.MOVE_AND_RESIZE;


public class Main {

    public static void main(String[] args) {
        try (FileOutputStream fos = new FileOutputStream("./workbook.xlsx");
             FileInputStream fis = new FileInputStream("./src/main/resources/orange.jpg")) {
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("sheet 1");
            byte[] byteOfImage = IOUtils.toByteArray(fis);
            int pictureIdx = wb.addPicture(byteOfImage, Workbook.PICTURE_TYPE_JPEG);

            Drawing drawing = sheet.createDrawingPatriarch();

            CreationHelper helper = wb.getCreationHelper();
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setAnchorType(MOVE_AND_RESIZE);

            anchor.setCol1(1);
            anchor.setCol2(4);

            anchor.setRow1(1);
            anchor.setRow2(4);

            anchor.setDx1(100);
            anchor.setDy1(100);

            anchor.setDx2(-100);
            anchor.setDy2(-100);

            Picture pict = drawing.createPicture(anchor, pictureIdx);

            wb.write(fos);
        } catch (FileNotFoundException fne) {
            System.out.println(fne.getMessage());
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

}
