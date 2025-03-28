module com.example.excelwordjava {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.web;

    requires java.desktop;

    requires org.controlsfx.controls;
    requires com.dlsc.formsfx;
    requires net.synedra.validatorfx;
    requires org.kordamp.ikonli.javafx;
    requires org.kordamp.bootstrapfx.core;
    requires eu.hansolo.tilesfx;
    requires com.almasb.fxgl.all;

    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires org.apache.poi.scratchpad;

    opens com.example.exel_word to javafx.fxml;
    exports com.example.exel_word;
}