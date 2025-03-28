package com.example.exel_word;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class FileMergerService {

    public void mergeToExcel(List<File> files, File outputFile) throws IOException {
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Объединенные данные");
        int rowNum = 0;

        for (File file : files) {
            try (Workbook workbook = WorkbookFactory.create(file)) {
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    for (Row row : sheet) {
                        Row newRow = outputSheet.createRow(rowNum++);
                        for (Cell cell : row) {
                            Cell newCell = newRow.createCell(cell.getColumnIndex(), cell.getCellType());
                            switch (cell.getCellType()) {
                                case STRING -> newCell.setCellValue(cell.getStringCellValue());
                                case NUMERIC -> newCell.setCellValue(cell.getNumericCellValue());
                                case BOOLEAN -> newCell.setCellValue(cell.getBooleanCellValue());
                                case FORMULA -> newCell.setCellFormula(cell.getCellFormula());
                                default -> newCell.setCellValue("");
                            }
                        }
                    }
                }
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
            outputWorkbook.write(fileOut);
        }
        outputWorkbook.close();
    }

    public void mergeToWord(List<File> files, File outputFile) throws IOException {
        XWPFDocument document = new XWPFDocument();

        for (File file : files) {
            addFileTitle(document, file.getName());

            try (Workbook workbook = WorkbookFactory.create(file)) {
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    XWPFTable table = document.createTable();
                    createTableFromSheet(sheet, table);
                    document.createParagraph().createRun().addBreak();
                }
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
            document.write(fileOut);
        }
        document.close();
    }

    private void addFileTitle(XWPFDocument document, String fileName) {
        XWPFParagraph title = document.createParagraph();
        XWPFRun titleRun = title.createRun();
        titleRun.setText(fileName);
        titleRun.setBold(true);
        titleRun.addBreak();
    }

    private void createTableFromSheet(Sheet sheet, XWPFTable table) {
        // Create headers
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                if (j >= table.getRow(0).getTableCells().size()) {
                    table.getRow(0).addNewTableCell();
                }
                Cell cell = headerRow.getCell(j);
                String value = (cell != null) ? FileHelper.getCellValueAsString(cell) : "";
                table.getRow(0).getCell(j).setText(value);
            }
        }

        // Create data rows
        for (int k = 1; k <= sheet.getLastRowNum(); k++) {
            Row row = sheet.getRow(k);
            if (row != null) {
                XWPFTableRow tableRow = table.createRow();
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    String value = (cell != null) ? FileHelper.getCellValueAsString(cell) : "";
                    if (j < tableRow.getTableCells().size()) {
                        tableRow.getCell(j).setText(value);
                    } else {
                        tableRow.addNewTableCell().setText(value);
                    }
                }
            }
        }
    }
}