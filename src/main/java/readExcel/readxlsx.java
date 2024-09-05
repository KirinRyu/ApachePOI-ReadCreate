package readExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class readxlsx {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\dan_r\\Project\\APACHE-POI\\teste\\teste.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(filePath);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = xssfWorkbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.println(cell.getStringCellValue() + "\t");
                            break;

                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue() + "\t");
                            break;

                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue() + "\t");
                            break;

                        default:
                            System.out.println("Tipo de célula desconhecida");
                            break;
                    }
                }
                System.out.println();
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
