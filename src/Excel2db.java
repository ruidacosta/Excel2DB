import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;

/**
 * Created by rui.alves.costa on 02-07-2015.
 */


public class Excel2db {
    public static void main(String[] args) {
        try {
            Workbook wb = WorkbookFactory.create(new File("C:\\Users\\rui.costa\\Desktop\\BPI_trades\\BPI_trades.xlsx"));
            Sheet sheet = wb.getSheetAt(0);
            for (Row row : sheet){
                for (Cell cell : row){
                    System.out.println(cell.getStringCellValue());
                }
                System.out.println("--------------------------------------");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }
}
