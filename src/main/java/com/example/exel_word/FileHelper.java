package com.example.exel_word;

import javafx.stage.FileChooser;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;
import org.apache.poi.ss.usermodel.Cell;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class FileHelper {

    // Выбор нескольких Excel файлов через диалог
    public static List<File> chooseExcelFiles(Window ownerWindow) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Excel Files", "*.xls", "*.xlsx"));
        return new ArrayList<>(fileChooser.showOpenMultipleDialog(ownerWindow));
    }

    // Выбор всех Excel файлов из папки
    public static List<File> chooseExcelFilesFromFolder(Window ownerWindow) {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        File folder = directoryChooser.showDialog(ownerWindow);
        List<File> files = new ArrayList<>();

        if (folder != null) {
            File[] foundFiles = folder.listFiles((dir, name) -> name.endsWith(".xls") || name.endsWith(".xlsx"));
            if (foundFiles != null) {
                files.addAll(List.of(foundFiles));
            }
        }
        return files;
    }

    // Выбор места сохранения (Excel/Word)
    public static File chooseOutputFile(Window ownerWindow, String format) {
        FileChooser fileChooser = new FileChooser();
        if (format.equals("Excel (.xlsx)")) {
            fileChooser.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
            fileChooser.setInitialFileName("merged.xlsx");
        } else {
            fileChooser.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Word Files", "*.docx"));
            fileChooser.setInitialFileName("merged.docx");
        }
        return fileChooser.showSaveDialog(ownerWindow);
    }

    // Просмотр файла в программе по умолчанию
    public static void previewFile(Window ownerWindow) {
        FileChooser fileChooser = new FileChooser();
        File file = fileChooser.showOpenDialog(ownerWindow);
        if (file != null) {
            try {
                java.awt.Desktop.getDesktop().open(file);
            } catch (IOException ex) {
                throw new RuntimeException("Не удалось открыть файл: " + ex.getMessage(), ex);
            }
        }
    }

    // Получение значения ячейки как строки (с обработкой типов данных)
    public static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            default -> "";
        };
    }
}