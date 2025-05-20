package com.example.exel_word;

import javafx.application.Application;

public class Main {
    public static void main(String[] args) {
        // Можно добавить предварительную логику
        System.out.println("Starting application...");

        // Запуск JavaFX
        Application.launch(HelloApplication.class, args);

        // Можно добавить логику после закрытия приложения
        System.out.println("Application closed");
    }
}