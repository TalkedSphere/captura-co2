/**
 * @file   CapturaApplication.java
 * @brief  Inicializa o projeto e implementa as regras de negócios.
 */

// Pacote.
package br.ufmg.capturaco2;

// Imports.
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;
import java.io.IOException;

// Classe pública.
public class CapturaApplication extends Application {
    /**
     * Inicia a tela e a mostra na tela.
     * @param stage Contâiner que representa a janela principal da aplicação.
     */
    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(CapturaApplication.class.getResource("CapturaView.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 714, 374); // Largura e altura da janela, respectivamente.
        stage.setTitle("Captura CO2");
        stage.setScene(scene);
        stage.show();
    }

    /**
     * Aplica as chamadas necessárias para o devido fechamento a janela.
     */
    @Override
    public void stop() {}

    /**
     * Main.
     */
    public static void main(String[] args) {
        launch();
    }
}