module br.ufmg.capturaco2 {
    requires javafx.controls;
    requires javafx.fxml;

    requires org.controlsfx.controls;
    requires org.kordamp.bootstrapfx.core;

    opens br.ufmg.capturaco2 to javafx.fxml;
    exports br.ufmg.capturaco2;
}