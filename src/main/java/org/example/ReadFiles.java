package org.example;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;


public class ReadFiles
{

    public static void main(String[] args) throws IOException {
        Path Otchet = null;
        Path all = null;
        Path path2 = Paths.get("C:\\Users\\sunri\\OneDrive\\Рабочий стол\\all.txt");
        if (!Files.exists(path2)) {
            all = Files.createFile(path2);
            System.out.println("Файл успешно создан");
        } else {
            System.out.println("Файл уже существует");
        }
        Scanner scanner = new Scanner(System.in);
        String fileName = scanner.nextLine();
        String fileName2 = scanner.nextLine();
        String fileName3 = scanner.nextLine();
        ExcelToTxt(fileName, fileName2, fileName3, all);
        NewExcelFile(all, Otchet);
        System.out.println("Запустить макрос?");
        if (scanner.nextLine().equals("yes"))
        {
            RunMacro(Otchet);
        }
    }


    public static void ExcelToTxt(String fileName, String fileName2, String fileName3, Path all) throws IOException
    {

        int c = 0;
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(fileName), StandardCharsets.UTF_16));
             BufferedReader reader2 = new BufferedReader(new InputStreamReader(new FileInputStream(fileName2), StandardCharsets.UTF_16));
             BufferedReader reader3 = new BufferedReader(new InputStreamReader(new FileInputStream(fileName3), StandardCharsets.UTF_16))
             ; BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(String.valueOf(all)), StandardCharsets.UTF_16))) {

            while (reader.ready())
            {
                String line = reader.readLine();
                String[] arrLine = line.split(",");
                c += 1;
                if (c != 1)
                {
                    writer.write(arrLine[0]);
                    writer.write(arrLine[1]);
                    writer.newLine();

                }
            }

            c = 0;
            while (reader2.ready())
            {
                String line = reader2.readLine();
                String[] arrLine = line.split(",");
                c += 1;
                if (c != 1)
                {
                    writer.write(arrLine[0]);
                    writer.write(arrLine[1]);
                    writer.newLine();
                }
            }
            c = 0;
            while (reader3.ready())
            {
                String line = reader3.readLine();
                String[] arrLine = line.split(",");
                c += 1;
                if (c != 1)
                {
                    writer.write(arrLine[0]);
                    writer.write(arrLine[1]);
                    writer.newLine();
                }
            }



            System.out.println("CSV успешно сконвертирован в текстовый файл с кодировкой Windows-1251.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void NewExcelFile(Path all, Path Otchet) throws IOException {
        Path path = Paths.get("C:\\Users\\sunri\\OneDrive\\Рабочий стол\\Otchet.xlsx");
        if (!Files.exists(path)) {
            Otchet = Files.createFile(path);
            System.out.println("Файл успешно создан");
        } else {
            System.out.println("Файл уже существует");
        }

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(String.valueOf(all)), StandardCharsets.UTF_16))) {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Data");

            int rowNum = 0;
            while (reader.ready()) {
                Row row = sheet.createRow(rowNum++);
                String line = reader.readLine();
                String[] values = line.split(","); // предполагается, что данные разделены запятыми
                for (int i = 0; i < values.length; i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(values[i]);
                }
            }
            String fileExtension = "xlsx";
            FileOutputStream fos = new FileOutputStream(String.valueOf(Otchet));
            workbook.write(fos);

            fos.close();
            workbook.close();

            System.out.println("Данные успешно записаны в Excel файл.");
        }
    }
    public static void RunMacro(Path Otchet)
    {
        try {
            FileInputStream fis = new FileInputStream(String.valueOf(Otchet));
            Workbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = (XSSFSheet) wb.getSheetAt(0);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Sub МакросДляВременииСтолбцов()\n" +
                    "' МакросДЛяВременииСтолбцов Макрос\n" +
                    "End Sub");

            // Запуск макроса
            sheet.getClass().getMethod("setForceFormulaRecalculation", boolean.class).invoke(sheet, true);

            // Сохраняем изменения
            FileOutputStream fileOut = new FileOutputStream(String.valueOf(Otchet));
            wb.write(fileOut);
            fileOut.close();

            System.out.println("Макрос успешно выполнен в Excel файле.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void CreateApp()
    {

    }
}