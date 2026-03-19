package org.example;

import org.json.JSONArray;
import org.json.JSONObject;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.stream.Collectors;

/**
 * Aplicação com interface gráfica para gerar documentos de internação.
 */
public class GeradorDocsInternacao extends JFrame {

    private static final String SERVER_URL = "https://teste.voticor.online";
    private final DefaultListModel<Paciente> listModel = new DefaultListModel<>();
    private final JList<Paciente> patientList = new JList<>(listModel);
    private final JButton generateButton = new JButton("Gerar Documentos...");
    private final JLabel statusLabel = new JLabel("Pronto.");
    private final List<Paciente> pacientes = new ArrayList<>();

    // Classe interna para armazenar os dados do paciente de forma organizada.
    private static class Paciente {
        String nome;
        String procedencia;
        String arquivo;

        Paciente(String nome, String procedencia, String arquivo) {
            this.nome = nome;
            this.procedencia = procedencia;
            this.arquivo = arquivo;
        }

        @Override
        public String toString() {
            return this.nome; // O texto que aparecerá na lista.
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                System.err.println("Não foi possível definir o Look and Feel do sistema.");
            }
            new GeradorDocsInternacao().setVisible(true);
        });
    }

    public GeradorDocsInternacao() {
        super("Gerador de Documentos de Internação");
        setupUI();
        refreshPatientList(); // Carrega a lista ao iniciar.
    }

    private void setupUI() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(600, 700);
        setLocationRelativeTo(null);

        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // Painel superior com título e botão de atualizar
        JPanel topPanel = new JPanel(new BorderLayout());
        topPanel.add(new JLabel("Selecione um paciente da lista abaixo:"), BorderLayout.NORTH);
        JButton refreshButton = new JButton("Atualizar Lista");
        refreshButton.addActionListener(e -> refreshPatientList());
        topPanel.add(refreshButton, BorderLayout.EAST);
        mainPanel.add(topPanel, BorderLayout.NORTH);

        // Lista de pacientes
        patientList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        patientList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                generateButton.setEnabled(patientList.getSelectedIndex() != -1);
            }
        });
        mainPanel.add(new JScrollPane(patientList), BorderLayout.CENTER);

        // Painel inferior com botão de gerar e status
        JPanel bottomPanel = new JPanel(new BorderLayout());
        generateButton.setEnabled(false);
        generateButton.addActionListener(e -> openDocumentsDialog());
        bottomPanel.add(generateButton, BorderLayout.CENTER);
        statusLabel.setBorder(BorderFactory.createEmptyBorder(5, 5, 0, 0));
        bottomPanel.add(statusLabel, BorderLayout.SOUTH);
        mainPanel.add(bottomPanel, BorderLayout.SOUTH);

        add(mainPanel);
    }

    private void refreshPatientList() {
        setStatus("Buscando lista de pacientes do dia...");
        new SwingWorker<List<Paciente>, Void>() {
            @Override
            protected List<Paciente> doInBackground() throws Exception {
                List<Paciente> result = new ArrayList<>();
                String json = httpGet(SERVER_URL + "/api/internacao/listar");
                JSONArray array = new JSONArray(json);
                for (int i = 0; i < array.length(); i++) {
                    JSONObject obj = array.getJSONObject(i);
                    result.add(new Paciente(
                            obj.getString("nome"),
                            obj.optString("procedencia", "Não informada"),
                            obj.getString("arquivo")
                    ));
                }
                return result;
            }

            @Override
            protected void done() {
                try {
                    pacientes.clear();
                    pacientes.addAll(get());
                    listModel.clear();
                    for (Paciente p : pacientes) {
                        listModel.addElement(p);
                    }
                    setStatus(pacientes.size() + " paciente(s) encontrado(s) para hoje.");
                } catch (Exception e) {
                    e.printStackTrace();
                    setStatus("Erro ao buscar pacientes.");
                    JOptionPane.showMessageDialog(GeradorDocsInternacao.this, "Erro ao buscar lista de pacientes: " + e.getMessage(), "Erro de Conexão", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
    }

    private void openDocumentsDialog() {
        Paciente selectedPatient = patientList.getSelectedValue();
        if (selectedPatient == null) {
            return;
        }

        // O restante da lógica de diálogo é a mesma que já tínhamos.
        mostrarDialogoOpcoes(selectedPatient);
    }

    private void mostrarDialogoOpcoes(Paciente paciente) {
        String nomePaciente = paciente.nome;
        String procedencia = paciente.procedencia;

        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JCheckBox chkReceita = new JCheckBox("Receita AAS + Clopidogrel", true);
        JCheckBox chkInternacao = new JCheckBox("Solicitação de Internação", true);
        JCheckBox chkSolAngio = new JCheckBox("Solicitação de Angioplastia (SUS)", true);
        JCheckBox chkJustAngio = new JCheckBox("Justificativa para Angioplastia", true);
        JCheckBox chkEvolucao = new JCheckBox("Evolução Tasy", true);

        JPanel generalPanel = new JPanel(new GridLayout(0, 2, 5, 5));
        generalPanel.setBorder(BorderFactory.createTitledBorder("Dados Gerais (Solicitação de Internação)"));
        generalPanel.add(new JLabel("Artérias Tratadas:"));
        JTextField arteriasField = new JTextField(20);
        generalPanel.add(arteriasField);
        generalPanel.add(new JLabel("Quantidade de Stents:"));
        JTextField stentsField = new JTextField(5);
        generalPanel.add(stentsField);

        JPanel justificativaOptionsPanel = new JPanel(new GridLayout(0, 2, 5, 5));
        justificativaOptionsPanel.setBorder(BorderFactory.createTitledBorder("Dados Específicos (Justificativa de Angioplastia)"));
        justificativaOptionsPanel.add(new JLabel("Artérias (Justificativa):"));
        JTextField arteriasJustField = new JTextField(20);
        justificativaOptionsPanel.add(arteriasJustField);
        justificativaOptionsPanel.add(new JLabel("Stents (Justificativa):"));
        JTextField stentsJustField = new JTextField(5);
        justificativaOptionsPanel.add(stentsJustField);
        justificativaOptionsPanel.setVisible(chkJustAngio.isSelected());
        chkJustAngio.addActionListener(e -> justificativaOptionsPanel.setVisible(chkJustAngio.isSelected()));

        JCheckBox chkConduta = new JCheckBox("Sugestão de Conduta");
        JPanel condutaOptionsPanel = new JPanel();
        condutaOptionsPanel.setLayout(new BoxLayout(condutaOptionsPanel, BoxLayout.Y_AXIS));
        condutaOptionsPanel.setBorder(BorderFactory.createTitledBorder("Opções de Conduta"));
        ButtonGroup condutaGroup = new ButtonGroup();
        JRadioButton rbClinico = new JRadioButton("Tratamento Clínico");
        JRadioButton rbAngio = new JRadioButton("Angioplastia Coronariana");
        JRadioButton rbCirurgia = new JRadioButton("Avaliação da cirurgia cardíaca");
        rbClinico.setSelected(true);
        condutaGroup.add(rbClinico);
        condutaGroup.add(rbAngio);
        condutaGroup.add(rbCirurgia);
        JTextArea obsTextArea = new JTextArea(3, 20);
        obsTextArea.setLineWrap(true);
        obsTextArea.setWrapStyleWord(true);
        JScrollPane obsScrollPane = new JScrollPane(obsTextArea);
        obsScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        condutaOptionsPanel.add(rbClinico);
        condutaOptionsPanel.add(rbAngio);
        condutaOptionsPanel.add(rbCirurgia);
        condutaOptionsPanel.add(Box.createVerticalStrut(10));
        condutaOptionsPanel.add(new JLabel("Observação:"));
        condutaOptionsPanel.add(obsScrollPane);
        condutaOptionsPanel.setVisible(false);
        chkConduta.addActionListener(e -> condutaOptionsPanel.setVisible(chkConduta.isSelected()));

        panel.add(chkReceita);
        panel.add(chkInternacao);
        panel.add(chkSolAngio);
        panel.add(chkJustAngio);
        panel.add(chkEvolucao);
        panel.add(Box.createVerticalStrut(10));
        panel.add(generalPanel);
        panel.add(justificativaOptionsPanel);
        panel.add(Box.createVerticalStrut(10));
        panel.add(chkConduta);
        panel.add(condutaOptionsPanel);

        JScrollPane mainScrollPane = new JScrollPane(panel);
        mainScrollPane.setPreferredSize(new Dimension(500, 600));
        mainScrollPane.setBorder(null);

        int result = JOptionPane.showConfirmDialog(this, mainScrollPane, "Gerar Documentos para: " + nomePaciente, JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        if (result == JOptionPane.OK_OPTION) {
            List<String> modelos = new ArrayList<>();
            if (chkReceita.isSelected()) modelos.add("receita_aas_clopidogrel.docx");
            if (chkInternacao.isSelected()) modelos.add("internacao.docx");
            if (chkSolAngio.isSelected()) modelos.add("solicitacao_angio_sus.docx");
            if (chkJustAngio.isSelected()) modelos.add("justificativa_angio.docx");
            if (chkEvolucao.isSelected()) modelos.add("evolucao_tasy.docx");

            Map<String, String> payloadData = new HashMap<>();
            payloadData.put("procedencia", procedencia);
            payloadData.put("arterias", arteriasField.getText());
            payloadData.put("stents", stentsField.getText());

            if (chkJustAngio.isSelected()) {
                payloadData.put("arterias_just", arteriasJustField.getText());
                payloadData.put("stents_just", stentsJustField.getText());
            }

            if (chkConduta.isSelected()) {
                modelos.add("sugestao_de_conduta.docx");
                payloadData.put("chk_clinico", rbClinico.isSelected() ? "X" : " ");
                payloadData.put("chk_angio", rbAngio.isSelected() ? "X" : " ");
                payloadData.put("chk_cirurgia", rbCirurgia.isSelected() ? "X" : " ");
                payloadData.put("obs_txt", obsTextArea.getText() != null ? obsTextArea.getText() : "");
            }

            gerarEImprimirDocumentos(nomePaciente, modelos, payloadData);
        }
    }

    private void gerarEImprimirDocumentos(String nome, List<String> modelos, Map<String, String> data) {
        if (modelos.isEmpty()) {
            setStatus("Nenhum modelo selecionado.");
            return;
        }

        setStatus("Gerando " + modelos.size() + " documento(s)...");
        new SwingWorker<List<File>, Void>() {
            @Override
            protected List<File> doInBackground() throws Exception {
                List<File> generatedFiles = new ArrayList<>();
                URL url = new URL(SERVER_URL + "/api/internacao/gerar");
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("POST");
                conn.setRequestProperty("Content-Type", "application/json");
                conn.setRequestProperty("Accept", "application/json");
                conn.setDoOutput(true);

                JSONObject payload = new JSONObject();
                payload.put("nome", nome);
                payload.put("modelos", new JSONArray(modelos));
                data.forEach(payload::put);

                try (OutputStream os = conn.getOutputStream()) {
                    os.write(payload.toString().getBytes(StandardCharsets.UTF_8));
                }

                if (conn.getResponseCode() != 201) {
                    throw new IOException("Erro no servidor: " + conn.getResponseCode());
                }

                try (BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
                    JSONObject response = new JSONObject(br.lines().collect(Collectors.joining("\n")));
                    JSONArray arquivosGerados = response.getJSONArray("arquivos_gerados");

                    for (int i = 0; i < arquivosGerados.length(); i++) {
                        String filename = arquivosGerados.getString(i);
                        File tempFile = File.createTempFile("internacao_", ".docx");
                        tempFile.deleteOnExit();
                        downloadFile(SERVER_URL + "/api/baixar/" + encodePath(filename), tempFile);
                        generatedFiles.add(tempFile);
                    }
                }
                return generatedFiles;
            }

            @Override
            protected void done() {
                try {
                    List<File> files = get();
                    setStatus(files.size() + " documento(s) gerado(s) com sucesso.");
                    for (File f : files) {
                        abrirArquivo(f);
                    }
                    JOptionPane.showMessageDialog(GeradorDocsInternacao.this, files.size() + " documento(s) abertos para impressão.", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception e) {
                    e.printStackTrace();
                    setStatus("Falha ao gerar documentos.");
                    JOptionPane.showMessageDialog(GeradorDocsInternacao.this, "Falha ao gerar/imprimir documentos: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
    }

    // --- MÉTODOS UTILITÁRIOS ---

    private void setStatus(String message) {
        statusLabel.setText(message);
    }

    private String httpGet(String urlStr) throws IOException {
        URL url = new URL(urlStr);
        HttpURLConnection http = (HttpURLConnection) url.openConnection();
        http.setRequestProperty("User-Agent", "GeradorDocs/1.0");
        http.setRequestProperty("Accept", "application/json");
        http.setConnectTimeout(10000);

        if (http.getResponseCode() >= 400) {
            throw new IOException("Erro do servidor: " + http.getResponseCode());
        }

        try (Scanner sc = new Scanner(http.getInputStream(), StandardCharsets.UTF_8.name())) {
            sc.useDelimiter("\\A");
            return sc.hasNext() ? sc.next() : "";
        }
    }

    private void downloadFile(String urlStr, File dest) throws IOException {
        URL url = new URL(urlStr);
        HttpURLConnection http = (HttpURLConnection) url.openConnection();
        http.setRequestProperty("User-Agent", "GeradorDocs/1.0");
        http.setConnectTimeout(15000);

        if (http.getResponseCode() != 200) {
            throw new IOException("Falha no download (" + http.getResponseCode() + "): " + urlStr);
        }

        try (InputStream in = http.getInputStream()) {
            Files.copy(in, dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }
    }

    private String encodePath(String filename) {
        try {
            return URLEncoder.encode(filename, StandardCharsets.UTF_8.name()).replace("+", "%20");
        } catch (UnsupportedEncodingException e) {
            return filename.replace(" ", "%20"); // Fallback
        }
    }

    private void abrirArquivo(File arquivo) {
        if (!Desktop.isDesktopSupported() || arquivo == null) {
            setStatus("Erro: Não é possível abrir arquivos neste sistema.");
            return;
        }
        new Thread(() -> {
            try {
                Thread.sleep(500); // Pequena pausa para o arquivo ser salvo
                Desktop.getDesktop().open(arquivo);
            } catch (Exception e) {
                setStatus("Erro ao abrir " + arquivo.getName());
                System.err.println("Falha ao abrir o arquivo: " + e.getMessage());
            }
        }).start();
    }
}