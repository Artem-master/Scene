
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
import java.util.Iterator;

public class test {
    public static void main(String[] args) throws IOException {
//        String dowPackPath = "C:\\Users\\go\\Desktop\\Сити Транс\\";
//        File file = new File(dowPackPath);
//        String [] arr = file.list();
//        Desktop desktop = Desktop.getDesktop();
//        for (String s : arr) {
//            File newFile = new File(dowPackPath + s);
//            desktop.print(newFile);
//        }
        String dow = "C:\\Users\\go\\Desktop\\Tarkett\\Получатели.xlsx";
        FileInputStream fis = new FileInputStream(dow);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sh = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sh.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            System.out.println(row.getCell(0).getStringCellValue());
        }
    }
}
