/**
 * @file   CapturaController.java
 * @brief  Gerencia os eventos da tela associando ações aos elementos.
 */

// Pacote.
package br.ufmg.capturaco2;

// Imports.
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.stage.Stage;

import java.io.*;

// Classe pública.
public class CapturaInicialController {
    // A tag @FXML indica um elemento pego diretamente do view.
    @FXML
    private Button botao_in_out;
    @FXML
    private Button botao_out;

    /**
     * Automaticamente executado após o carregamento do view.
     */
    @FXML
    public void initialize() {
        // Define as ações do botão reponsável mostrar a janela da análise IN OUT.
        botao_in_out.setOnAction(event -> {
            try {
                Parent root = FXMLLoader.load(getClass().getResource("CapturaInOutView.fxml"));
                Stage stage = (Stage) ((Node) event.getSource()).getScene().getWindow();

                Scene scene = new Scene(root, 714, 374);
                stage.setScene(scene);
                stage.show();
            } catch (IOException _) {}
        });

        // Define as ações do botão reponsável mostrar a janela da análise OUT.
        botao_out.setOnAction(event -> {
            try {
                Parent root = FXMLLoader.load(getClass().getResource("CapturaOutView.fxml"));
                Stage stage = (Stage) ((Node) event.getSource()).getScene().getWindow();

                Scene scene = new Scene(root, 714, 374);
                stage.setScene(scene);
                stage.show();
            } catch (IOException _) {}
        });
    }
}