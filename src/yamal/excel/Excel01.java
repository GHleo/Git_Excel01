package yamal.excel;

import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel01 {
    public static void main(String[] args) throws Exception{
        try {
            FileInputStream file = new FileInputStream(new File("//Users//ileo//IdeaProjects//test.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator(); // Iterator для перебора строк
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()){
                        case NUMERIC:
                            System.out.printf("%.0f", cell.getNumericCellValue());
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue()+"\t\t");
                            break;

                    }
                }
                System.out.println();
            }
            file.close();
        }
        catch (Exception e){
            System.out.println("whats wrong!");

        }
    }

}
