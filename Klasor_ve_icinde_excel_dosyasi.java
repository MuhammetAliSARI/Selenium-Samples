package test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Klasor_ve_icinde_excel_dosyasi{
    public static void main(String[] args){

// EXEL İÇİN KLASÖR
        String folderPath = "C:/EXCEL-JAVA";
        File folder = new File(folderPath);

        if (!folder.exists()) {
            boolean isCreated = folder.mkdirs();

            if (isCreated) {
                System.out.println("Klasör oluşturuldu.");
            } else {
                System.out.println("Klasör oluşturulamadı.");
            }
        } else {
            System.out.println("Klasör zaten mevcut.");
        }

        Workbook workbook = new XSSFWorkbook();                // Yeni bir XSSFWorkbook nesnesi oluşturun
        Sheet sheet = workbook.createSheet("Sayfa-1");      // Yeni bir sayfa oluşturun
// EXCEL DOSYASI OLUŞTURMA
            // İlk satırı oluşturup ve başlıkları ekliypruz
            Row headerRow = sheet.createRow(0);
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("1.SÜTUN");
            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("2.SÜTUN");
            Cell headerCell3 = headerRow.createCell(2);
            headerCell3.setCellValue("3.SÜTUN");

            // Verileri eklemek için yeni satırlar oluşturuyoruz
            Row dataRow1 = sheet.createRow(1);
            Cell dataCell1 = dataRow1.createCell(0);
            dataCell1.setCellValue("1.sütun verisi");
            Cell dataCell2 = dataRow1.createCell(1);
            dataCell2.setCellValue("2.sütun verisi");
            Cell dataCell3 = dataRow1.createCell(2);
            dataCell3.setCellValue("2.sütun verisi");

        // Dosyayı kaydetmek için bir FileOutputStream kullanıyouz.
        try (FileOutputStream fileOutputStream = new FileOutputStream("C:/EXCEL-JAVA/ExcelJava.xlsx")) {
            workbook.write(fileOutputStream);
            System.out.println("XLSX dosyası oluşturuldu.");
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Workbook ve FileOutputStream'i kapatıyoruz
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
