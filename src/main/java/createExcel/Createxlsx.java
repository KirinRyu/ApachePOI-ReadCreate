package createExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Createxlsx {

    public static void main(String[] args) {

        String apresentacao0 = "Olá ";
        String apresentacao1 = "Mundo!";

        try(XSSFWorkbook xssfWorkbook = new XSSFWorkbook()) {
            Sheet sheet = xssfWorkbook.createSheet("Teste");

            Row row0 = sheet.createRow(0);

            Cell cell0 = row0.createCell(0);
            cell0.setCellValue(apresentacao0);

            Cell cell1 = row0.createCell(1);
            cell1.setCellValue(apresentacao1);

            try(FileOutputStream fileOutputStream = new FileOutputStream("teste.xlsx")) {
                xssfWorkbook.write(fileOutputStream);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
