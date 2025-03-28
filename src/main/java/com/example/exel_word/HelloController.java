package com.example.exel_word;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.DirectoryChooser;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.ArrayList;

public class HelloController {
    @FXML private ListView<String> fileListView;
    @FXML private ComboBox<String> outputFormatCombo;
    @FXML private Button addFilesButton;
    @FXML private Button addFolderButton;
    @FXML private Button removeButton;
    @FXML private Button mergeButton;
    @FXML private Button previewButton;

    private ObservableList<String> fileNames = FXCollections.observableArrayList();
    private List<File> selectedFiles = new ArrayList<>();
    private final FileMergerService fileMergerService = new FileMergerService();
    private final AlertHelper alertHelper = new AlertHelper();

    @FXML
    public void initialize() {
        fileListView.setItems(fileNames);
        outputFormatCombo.getItems().addAll("Excel (.xlsx)", "Word (.docx)");
        outputFormatCombo.getSelectionModel().selectFirst();
    }

    @FXML
    private void handleAddFiles() {
        List<File> files = FileHelper.chooseExcelFiles(addFilesButton.getScene().getWindow());
        if (files != null) {
            for (File file : files) {
                if (!selectedFiles.contains(file)) {
                    selectedFiles.add(file);
                    fileNames.add(file.getName());
                }
            }
        }
    }

    @FXML
    private void handleAddFolder() {
        List<File> files = FileHelper.chooseExcelFilesFromFolder(addFolderButton.getScene().getWindow());
        if (files != null) {
            for (File file : files) {
                if (!selectedFiles.contains(file)) {
                    selectedFiles.add(file);
                    fileNames.add(file.getName());
                }
            }
        }
    }

    @FXML
    private void handleRemove() {
        int selectedIndex = fileListView.getSelectionModel().getSelectedIndex();
        if (selectedIndex >= 0) {
            selectedFiles.remove(selectedIndex);
            fileNames.remove(selectedIndex);
        }
    }

    @FXML
    private void handleMerge() {
        if (selectedFiles.isEmpty()) {
            alertHelper.showAlert("Ошибка", "Нет файлов для объединения", Alert.AlertType.ERROR);
            return;
        }

        String format = outputFormatCombo.getValue();
        File outputFile = FileHelper.chooseOutputFile(mergeButton.getScene().getWindow(), format);

        if (outputFile != null) {
            try {
                if (format.equals("Excel (.xlsx)")) {
                    fileMergerService.mergeToExcel(selectedFiles, outputFile);
                } else {
                    fileMergerService.mergeToWord(selectedFiles, outputFile);
                }
                alertHelper.showAlert("Успех", "Файлы успешно объединены!", Alert.AlertType.INFORMATION);
            } catch (IOException ex) {
                alertHelper.showAlert("Ошибка", "Ошибка при объединении файлов: " + ex.getMessage(), Alert.AlertType.ERROR);
            }
        }
    }

    @FXML
    private void handlePreview() {
        FileHelper.previewFile(previewButton.getScene().getWindow());
    }
}