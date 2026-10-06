/**
 * @file   CapturaController.java
 * @brief  Gerencia os eventos da tela associando ações aos elementos.
 */

// Pacote.
package br.ufmg.capturaco2;

// Imports.
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

// Classe pública.
public class CapturaInOutController {
    // Constantes
    private static final String GRAFICO_EXPERIMENTAL = "0";
    private static final String GRAFICO_AJUSTADO = "1";
    private static final String GRAFICO_GERAL = "2";

    // A tag @FXML indica um elemento pego diretamente do view.
    @FXML
    private Label label_arquivo;
    @FXML
    private Button botao_arquivo;
    @FXML
    private TextField seg_inicial;
    @FXML
    private TextField seg_final;
    @FXML
    private Button botao_confirmar;
    @FXML
    private Button botao_grafico_experimental;
    @FXML
    private Button botao_grafico_ajustado;
    @FXML
    private Button botao_grafico_geral;
    @FXML
    private Button botao_voltar;

    // Variável que guarda o arquivo anexado.
    private File arquivo;
    private boolean transformar = false;
    private boolean confirmado = false;

    /**
     * Automaticamente executado após o carregamento do view.
     */
    @FXML
    public void initialize() {
        // Define as ações do botão reponsável por anexar o arquivo.
        botao_arquivo.setOnAction(event -> {
            // Criar o elemento de escolha do arquivo.
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Selecionar Arquivo");

            // Filtra as opções de arquivo.
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Planilhas", "*.xlsx", "*.xls", "*.csv"));

            // Abre a tela para a escolha do arquivo.
            Stage stage = (Stage) botao_arquivo.getScene().getWindow();
            arquivo = fileChooser.showOpenDialog(stage);

            // Garante que o arquivo seja válido.
            if (arquivo != null) {
                String nomeArquivo = arquivo.getName().toLowerCase();

                if (nomeArquivo.endsWith(".xlsx") || nomeArquivo.endsWith(".xls") || nomeArquivo.endsWith(".csv")) {
                    label_arquivo.setText(arquivo.getName());
                } else {
                    Alert alerta = new Alert(Alert.AlertType.ERROR);
                    alerta.setTitle("Arquivo Inválido");
                    alerta.setHeaderText("Formato de arquivo não suportado");
                    alerta.setContentText("Por favor, selecione uma planilha válida.");
                    alerta.showAndWait();
                    arquivo = null;
                }
            }

            // Roda os códigos necessários.
            if (arquivo != null) {
                new Thread(() -> transformador(arquivo.getName().toLowerCase())).start();
            }
        });

        // Define as ações do botão reponsável por confirmar a janela estável.
        botao_confirmar.setOnAction(event -> {
            if (arquivo == null && seg_inicial.getText() != null && seg_final.getText() != null ) {
                alertaArquivoJanela();
                return;
            }
            confirmado = true;
            new Thread(() -> processador_in_out()).start();
        });

        // Define as ações do botão responsável por gerar o gráfico experimental.
        botao_grafico_experimental.setOnAction(event -> {
            if (arquivo == null && confirmado) {
                alertaArquivoJanela();
                return;
            }
            new Thread(() -> geradorGrafico(GRAFICO_EXPERIMENTAL)).start();
        });

        // Define as ações do botão responsável por gerar o gráfico ajustado.
        botao_grafico_ajustado.setOnAction(event -> {
            if (arquivo == null && confirmado) {
                alertaArquivoJanela();
                return;
            }
            new Thread(() -> geradorGrafico(GRAFICO_AJUSTADO)).start();
        });

        // Define as ações do botão responsável por gerar o gráfico geral.
        botao_grafico_geral.setOnAction(event -> {
            if (arquivo == null && confirmado) {
                alertaArquivoJanela();
                return;
            }
            new Thread(() -> geradorGrafico(GRAFICO_GERAL)).start();
        });

        // Define as ações do botão reponsável mostrar a janela inicial
        botao_voltar.setOnAction(event -> {
            try {
                Parent root = FXMLLoader.load(getClass().getResource("CapturaInicialView.fxml"));
                Stage stage = (Stage) ((Node) event.getSource()).getScene().getWindow();

                Scene scene = new Scene(root, 714, 260);
                stage.setScene(scene);
                stage.show();
            } catch (IOException _) {}
        });
    }

    /**
     * Transforma um arquivo .csv ou .xls para .xlsx.
     */
    private void transformador(String nomeArquivo) {
        // Cria o arquivo temporário.
        File arquivoTemporario = null;

        // Se o arquivo for um .csv, muda para .xlsx.
        if(nomeArquivo.endsWith(".csv")) {
            // Executa o arquivo temporário.
            try {
                arquivoTemporario = criarArquivoTemp("converter_csv_para_xlsx");
                // Descobre o novo caminho substituindo .csv por .xlsx
                String caminhoXlsx = arquivo.getAbsolutePath().replaceAll("(?i)\\.csv$", ".xlsx");
                if(arquivoTemporario != null) {
                    transformar = true;
                    executarArquivoTemp(arquivoTemporario, caminhoXlsx, null);
                    transformar = false;
                }
            } catch (IOException e) {
                Platform.runLater(()  -> alertaErroSistema(e));
            }

            // Apaga o arquivo temporário.
            if (arquivoTemporario != null && arquivoTemporario.exists()) arquivoTemporario.delete();
        }

        // Se o arquivo for um .xls, muda para .xlsx.
        if(nomeArquivo.endsWith(".xls")) {
            // Executa o arquivo temporário.
            try {
                arquivoTemporario = criarArquivoTemp("converter_xls_para_xlsx");
                // Descobre o novo caminho substituindo .csv por .xlsx
                String caminhoXlsx = arquivo.getAbsolutePath().replaceAll("(?i)\\.xls$", ".xlsx");
                if(arquivoTemporario != null) {
                    transformar = true;
                    executarArquivoTemp(arquivoTemporario, caminhoXlsx, null);
                    transformar = false;
                }
            } catch (IOException e) {
                Platform.runLater(()  -> alertaErroSistema(e));
            }

            // Apaga o arquivo temporário.
            if (arquivoTemporario != null && arquivoTemporario.exists()) arquivoTemporario.delete();
        }
    }

    /**
     * Processa os dados do arquivo separando o IN e o OUT.
     */
    private void processador_in_out() {
        // Cria o arquivo temporário.
        File arquivoTemporario = null;
        // Janela estável.
        if(seg_inicial.getText() == null || seg_final.getText() == null) {
            alertaJanelaEstavel();
            return;
        }
        String segInicial = seg_inicial.getText();
        String segFinal = seg_final.getText();

        // Executa o arquivo temporário.
        try {
            arquivoTemporario = criarArquivoTemp("processador_co2_in_out");
            if(arquivoTemporario != null) executarArquivoTemp(arquivoTemporario, segInicial, segFinal);
        } catch (IOException e) {
            Platform.runLater(()  -> alertaErroSistema(e));
        }

        // Apaga o arquivo temporário.
        if (arquivoTemporario != null && arquivoTemporario.exists()) arquivoTemporario.delete();
    }

    /**
     * Gera os gráficos.
     */
    private void geradorGrafico(String tipo) {
        // Cria o arquivo temporário.
        File arquivoTemporario = null;

        // Executa o arquivo temporário.
        try {
            arquivoTemporario = criarArquivoTemp("gerador_graficos_in_out");
            if(arquivoTemporario != null) executarArquivoTemp(arquivoTemporario, tipo, null);
        } catch (IOException e) {
            Platform.runLater(()  -> alertaErroSistema(e));
        }

        // Apaga o arquivo temporário.
        if (arquivoTemporario != null && arquivoTemporario.exists()) arquivoTemporario.delete();
    }

    /**
     * Alerta de escolha de arquivo.
     */
    private File criarArquivoTemp(String nomeArquivo) {
        File arquivoTemporario = null;
        try {
            // Cria o arquivo temporário.
            arquivoTemporario = File.createTempFile(nomeArquivo, ".py");
            arquivoTemporario.deleteOnExit();
            // Copia o script para o arquivo temporário.
            try (InputStream is = getClass().getResourceAsStream("/" + nomeArquivo + ".py")) {
                Files.copy(is, arquivoTemporario.toPath(), StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (IOException e) {
            Platform.runLater(()  -> {
                Alert alerta = new Alert(Alert.AlertType.ERROR);
                alerta.setTitle("Erro no sistema!!");
                alerta.setContentText("Erro de I/O: " + e.getMessage());
                alerta.showAndWait();
            });
        }

        return arquivoTemporario;
    }

    /**
     * Executa um arquivo e lê a sua saída.
     */
    private void executarArquivoTemp(File arquivoTemporario, String informacaoAdicional1, String informacaoAdicional2) throws IOException {
        // Roda o script do arquivo temporário.
        List<String> comandos = new ArrayList<>(Arrays.asList(
            "python",
            arquivoTemporario.getAbsolutePath(),
            arquivo.getAbsolutePath()
        ));
        if (informacaoAdicional1 != null) comandos.add(informacaoAdicional1);
        if (informacaoAdicional2 != null) comandos.add(informacaoAdicional2);
        ProcessBuilder pb = new ProcessBuilder(comandos);
        pb.redirectErrorStream(true);
        Process processo = pb.start();

        // Lê a saída do console do python.
        StringBuilder output = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(processo.getInputStream()))) {
            String linha;
            while ((linha = reader.readLine()) != null) {
                System.out.println(linha);
                output.append(linha).append("\n");
            }
        }

        try {
            // Espera até que o script termine de ser executado e se der erro mostra-o na tela.
            int codigoSaida = processo.waitFor();
            if(codigoSaida == 0 && informacaoAdicional1 != null && transformar) arquivo = new File(informacaoAdicional1);
            else if (codigoSaida != 0) {
                String mensagemErro = output.toString();
                Platform.runLater(() -> {
                    Alert alerta = new Alert(Alert.AlertType.ERROR);
                    alerta.setTitle("Erro no arquivo: " + arquivoTemporario.getName());
                    alerta.setHeaderText("O script Python falhou (Código " + codigoSaida + ")");
                    alerta.setContentText(mensagemErro.isEmpty() ? "Verifique se o Python está instalado corretamente." : mensagemErro);
                    alerta.showAndWait();
                });
            }
        }  catch (InterruptedException e) {
            Platform.runLater(()  -> alertaErroSistema(e));
            Thread.currentThread().interrupt();
        }
    }

    /**
     * Alerta de escolha de arquivo.
     */
    private void alertaArquivoJanela() {
        Alert alerta = new Alert(Alert.AlertType.WARNING);
        alerta.setTitle("Atenção");
        alerta.setContentText("Selecione um arquivo primeiro ou defina a janela estável!!");
        alerta.showAndWait();
    }

    /**
     * Alerta da janela estável;
     */
    private void alertaJanelaEstavel() {
        Alert alerta = new Alert(Alert.AlertType.WARNING);
        alerta.setTitle("Atenção");
        alerta.setContentText("Defina a janela estável!!!");
        alerta.showAndWait();
    }

    /**
     * Alerta de erro no sistema.
     */
    private void alertaErroSistema(Exception e) {
        Alert alerta = new Alert(Alert.AlertType.ERROR);
        alerta.setTitle("Erro no sistema!!");
        alerta.setContentText("Erro: " + e.getMessage());
        alerta.showAndWait();
    }
}