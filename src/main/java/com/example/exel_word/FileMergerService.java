package com.example.exel_word;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class FileMergerService {

    public void mergeToExcel(List<File> files, File outputFile) throws IOException {
        // Создаем новую книгу Excel для результата
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Объединенные данные");
        int rowNum = 0;

        // Первый проход: собираем все уникальные заголовки
        Set<String> allHeaders = new LinkedHashSet<>();
        List<Map<String, String>> allData = new ArrayList<>();

        for (File file : files) {
            try (Workbook workbook = WorkbookFactory.create(file)) {
                // Обрабатываем каждый лист
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    List<String> headers = new ArrayList<>();
                    Row headerRow = sheet.getRow(0);

                    // Собираем заголовки из первой строки
                    if (headerRow != null) {
                        for (Cell cell : headerRow) {
                            headers.add(cell.getStringCellValue());
                        }
                        allHeaders.addAll(headers);
                    }

                    // Собираем данные из строк
                    for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                        Row row = sheet.getRow(j);
                        if (row != null) {
                            Map<String, String> rowData = new HashMap<>();
                            for (int k = 0; k < headers.size(); k++) {
                                Cell cell = row.getCell(k);
                                // Преобразуем значение ячейки в строку
                                String value = (cell != null) ? FileHelper.getCellValueAsString(cell) : "";
                                rowData.put(headers.get(k), value);
                            }
                            allData.add(rowData);
                        }
                    }
                }
            }
        }

        // Записываем заголовки в результирующий файл
        Row headerRow = outputSheet.createRow(rowNum++);
        int colNum = 0;
        for (String header : allHeaders) {
            Cell cell = headerRow.createCell(colNum++);
            cell.setCellValue(header);
        }

        // Записываем данные
        for (Map<String, String> rowData : allData) {
            Row row = outputSheet.createRow(rowNum++);
            colNum = 0;
            for (String header : allHeaders) {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(rowData.getOrDefault(header, ""));
            }
        }

        // Сохраняем результат в файл
        try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
            outputWorkbook.write(fileOut);
        }
        outputWorkbook.close();
    }

    public void mergeToWord(List<File> files, File outputFile) throws IOException {
        // Создаем новый документ Word
        XWPFDocument document = new XWPFDocument();

        // Собираем все данные
        Set<String> allHeaders = new LinkedHashSet<>();
        List<Map<String, String>> allData = new ArrayList<>();

        for (File file : files) {
            try (Workbook workbook = WorkbookFactory.create(file)) {
                // Обрабатываем каждый лист
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    List<String> headers = new ArrayList<>();
                    Row headerRow = sheet.getRow(0);

                    // Собираем заголовки
                    if (headerRow != null) {
                        for (Cell cell : headerRow) {
                            headers.add(cell.getStringCellValue());
                        }
                        allHeaders.addAll(headers);
                    }

                    // Проверяем совместимость столбцов
                    if (!headers.isEmpty() && !allHeaders.isEmpty() && !headers.equals(new ArrayList<>(allHeaders))) {
                        throw new IOException("Файл " + file.getName() + " содержит несовместимые столбцы");
                    }

                    // Собираем данные
                    for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                        Row row = sheet.getRow(j);
                        if (row != null) {
                            Map<String, String> rowData = new HashMap<>();
                            for (int k = 0; k < headers.size(); k++) {
                                Cell cell = row.getCell(k);
                                String value = (cell != null) ? FileHelper.getCellValueAsString(cell) : "";
                                rowData.put(headers.get(k), value);
                            }
                            allData.add(rowData);
                        }
                    }
                }
            }
        }

        // Создаем объединенную таблицу
        XWPFTable table = document.createTable();

        // Добавляем заголовки
        XWPFTableRow headerTableRow = table.getRow(0);
        int colIndex = 0;
        for (String header : allHeaders) {
            if (colIndex >= headerTableRow.getTableCells().size()) {
                headerTableRow.addNewTableCell();
            }
            headerTableRow.getCell(colIndex).setText(header);
            colIndex++;
        }

        // Добавляем строки с данными
        for (Map<String, String> rowData : allData) {
            XWPFTableRow tableRow = table.createRow();
            colIndex = 0;
            for (String header : allHeaders) {
                if (colIndex >= tableRow.getTableCells().size()) {
                    tableRow.addNewTableCell();
                }
                tableRow.getCell(colIndex).setText(rowData.getOrDefault(header, ""));
                colIndex++;
            }
        }

        // Сохраняем документ
        try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
            document.write(fileOut);
        }
        document.close();
    }
}