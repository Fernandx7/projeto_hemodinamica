import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.json.JSONArray;
import org.json.JSONObject;


public class AppHemo extends JFrame {

    // --- COMPONENTES UI ---
    private JTabbedPane tabbedPane;
    private JTable tablePendentes, tableHistorico, tableRelatorio, tableMsgTasy, tableInternacao;
    private DefaultTableModel modelPendentes, modelHistorico, modelRelatorio, modelMsgTasy, modelInternacao;
    private JLabel statusLabel, infoLabel;
    private JButton btnAbrirUltimo;
    private File lastGeneratedFile;


    // --- CONFIGURAÇÕES DA APLICAÇÃO ---
    private static final String CONFIG_FILE = "config.json";
    private static final String CURRENT_VERSION = "9.7"; // Versão atualizada
    private String serverUrl = "https://teste.voticor.online";
    private String rootPath = "A:\\projetos atuais\\TRABALHOVOTICOR\\laudos";
    private String lastVersionSeen = "";


    public static void main(String[] args) {
        System.setProperty("https.protocols", "TLSv1.2");
        SwingUtilities.invokeLater(() -> new AppHemo().setVisible(true));
    }

    public AppHemo() {
        super("VOTICOR - Hemodinâmica v" + CURRENT_VERSION);
        loadConfig();
        setupUI();
        refreshAll();
    }

    private void setupUI() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 800);

        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (Exception e) {
            try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); } catch (Exception ignored) {}
        }

        JPanel topPanel = new JPanel(new BorderLayout(10, 10));
        topPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 5, 10));
        JLabel title = new JLabel("Sistema de Laudos");
        title.setFont(new Font("Arial", Font.BOLD, 20));
        title.setForeground(new Color(0, 51, 102));

        JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        btnAbrirUltimo = new JButton("Abrir Último Laudo");
        btnAbrirUltimo.setEnabled(false);
        btnAbrirUltimo.addActionListener(e -> {
            if (lastGeneratedFile != null) {
                abrirArquivo(lastGeneratedFile);
            } else {
                JOptionPane.showMessageDialog(this, "Nenhum laudo foi gerado nesta sessão ainda.", "Aviso", JOptionPane.WARNING_MESSAGE);
            }
        });

        JButton btnConfig = new JButton("⚙ Config");
        JButton btnRefresh = new JButton("↻ Atualizar Tudo");
        btnConfig.addActionListener(e -> showConfigDialog());
        btnRefresh.addActionListener(e -> refreshAll());
        btnPanel.add(btnAbrirUltimo);
        btnPanel.add(btnConfig);
        btnPanel.add(btnRefresh);

        topPanel.add(title, BorderLayout.WEST);
        topPanel.add(btnPanel, BorderLayout.EAST);

        infoLabel = new JLabel("Conectado a: " + serverUrl);
        infoLabel.setBorder(BorderFactory.createEmptyBorder(0, 15, 10, 0));
        infoLabel.setForeground(Color.GRAY);

        tabbedPane = new JTabbedPane();

        modelPendentes = new DefaultTableModel(new String[]{"Nome", "Arquivo", "Origem", "Procedimento"}, 0) {
            @Override public boolean isCellEditable(int r, int c) { return false; }
        };
        tablePendentes = new JTable(modelPendentes);
        tablePendentes.setRowHeight(30);
        tablePendentes.getColumnModel().getColumn(0).setPreferredWidth(250);
        tablePendentes.getColumnModel().getColumn(1).setPreferredWidth(250);
        setupPendentesRightClick();

        JPanel panelPend = new JPanel(new BorderLayout());
        panelPend.add(new JScrollPane(tablePendentes), BorderLayout.CENTER);

        JButton btnIniciar = new JButton("ABRIR LAUDO");
        btnIniciar.setFont(new Font("Arial", Font.BOLD, 14));
        btnIniciar.setBackground(new Color(0, 102, 204));
        btnIniciar.setForeground(Color.WHITE);
        btnIniciar.setOpaque(true);
        btnIniciar.addActionListener(e -> new Thread(this::fluxoInteligente).start());

        JButton btnInternado = new JButton("Puxar Internado");
        btnInternado.setFont(new Font("Arial", Font.BOLD, 12));
        btnInternado.addActionListener(e -> new Thread(this::fluxoInternado).start());

        JPanel panelBotoes = new JPanel(new GridBagLayout());
        panelBotoes.setPreferredSize(new Dimension(0, 50));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.gridy = 0;
        gbc.gridx = 0;
        gbc.weightx = 0.85;
        panelBotoes.add(btnIniciar, gbc);
        gbc.gridx = 1;
        gbc.weightx = 0.15;
        panelBotoes.add(btnInternado, gbc);
        panelPend.add(panelBotoes, BorderLayout.SOUTH);

        modelHistorico = new DefaultTableModel(new String[]{"Arquivo Processado"}, 0) {
            @Override public boolean isCellEditable(int r, int c) { return false; }
        };
        tableHistorico = new JTable(modelHistorico);
        tableHistorico.setRowHeight(25);
        tableHistorico.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        JPanel panelHist = new JPanel(new BorderLayout());
        panelHist.add(new JScrollPane(tableHistorico), BorderLayout.CENTER);
        JPanel pBotoesHist = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton btnAngioHist = new JButton("GERAR EVOLUÇÃO (ANGIOPLASTIA)");
        btnAngioHist.addActionListener(e -> new Thread(this::fluxoAngioplastiaHistorico).start());
        JButton btnExcluir = new JButton("EXCLUIR REGISTRO(S)");
        btnExcluir.setBackground(new Color(200, 50, 50));
        btnExcluir.setForeground(Color.WHITE);
        btnExcluir.setOpaque(true);
        btnExcluir.addActionListener(e -> excluirRegistrosHistorico());
        pBotoesHist.add(btnAngioHist);
        pBotoesHist.add(btnExcluir);
        panelHist.add(pBotoesHist, BorderLayout.SOUTH);

        modelRelatorio = new DefaultTableModel(new String[]{"Paciente", "Origem", "Procedimento", "Conclusão (Stents)"}, 0);
        tableRelatorio = new JTable(modelRelatorio);
        tableRelatorio.setRowHeight(30);
        tableRelatorio.getColumnModel().getColumn(0).setPreferredWidth(200);
        tableRelatorio.getColumnModel().getColumn(3).setPreferredWidth(400);
        JPanel panelRel = new JPanel(new BorderLayout());
        panelRel.add(new JScrollPane(tableRelatorio), BorderLayout.CENTER);
        JButton btnCarregarRel = new JButton("Atualizar Relatório do Dia");
        btnCarregarRel.addActionListener(e -> gerarRelatorioDiario());
        panelRel.add(btnCarregarRel, BorderLayout.NORTH);

        modelMsgTasy = new DefaultTableModel(new String[]{"Paciente", "Procedimento"}, 0);
        tableMsgTasy = new JTable(modelMsgTasy);
        tableMsgTasy.setRowHeight(30);
        JPanel panelTasy = new JPanel(new BorderLayout());
        panelTasy.add(new JScrollPane(tableMsgTasy), BorderLayout.CENTER);
        JButton btnCarregarTasy = new JButton("Atualizar Relatório do Dia");
        btnCarregarTasy.addActionListener(e -> gerarRelatorioDiario());
        panelTasy.add(btnCarregarTasy, BorderLayout.NORTH);

        // (Ponto 2) Adicionada coluna 'Procedencia' ao modelo da tabela
        modelInternacao = new DefaultTableModel(new String[]{"Nome", "Data", "Arquivo", "Procedencia"}, 0) {
            @Override public boolean isCellEditable(int r, int c) { return false; }
        };
        tableInternacao = new JTable(modelInternacao);
        tableInternacao.setRowHeight(30);
        tableInternacao.getColumnModel().getColumn(0).setPreferredWidth(300);
        // Ocultar a coluna de procedência, usada apenas internamente
        tableInternacao.getColumnModel().getColumn(3).setMinWidth(0);
        tableInternacao.getColumnModel().getColumn(3).setMaxWidth(0);
        tableInternacao.getColumnModel().getColumn(3).setWidth(0);

        JPanel panelInternacao = new JPanel(new BorderLayout());
        panelInternacao.add(new JScrollPane(tableInternacao), BorderLayout.CENTER);
        JPanel pBotoesInternacao = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton btnAtualizarInternacao = new JButton("Atualizar Lista de Internação");
        btnAtualizarInternacao.addActionListener(e -> refreshListInternacao());
        JButton btnGerarDocsInternacao = new JButton("Gerar Documentos de Internação");
        btnGerarDocsInternacao.setBackground(new Color(23, 162, 184));
        btnGerarDocsInternacao.setForeground(Color.WHITE);
        btnGerarDocsInternacao.addActionListener(e -> abrirDialogoGeracaoDocs());
        pBotoesInternacao.add(btnAtualizarInternacao);
        pBotoesInternacao.add(btnGerarDocsInternacao);
        panelInternacao.add(pBotoesInternacao, BorderLayout.SOUTH);

        tabbedPane.addTab("Pendentes", panelPend);
        tabbedPane.addTab("Histórico", panelHist);
        tabbedPane.addTab("Relatório Diário", panelRel);
        tabbedPane.addTab("Mensagem Tasy", panelTasy);
        tabbedPane.addTab("Internação", panelInternacao);

        statusLabel = new JLabel("Pronto.");
        statusLabel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        JPanel main = new JPanel(new BorderLayout());
        main.add(topPanel, BorderLayout.NORTH);
        JPanel centerWrapper = new JPanel(new BorderLayout());
        centerWrapper.add(infoLabel, BorderLayout.NORTH);
        centerWrapper.add(tabbedPane, BorderLayout.CENTER);
        main.add(centerWrapper, BorderLayout.CENTER);
        main.add(statusLabel, BorderLayout.SOUTH);
        add(main);
    }

    private void fluxoInternado() {
        try {
            int resp = JOptionPane.showOptionDialog(this, "O paciente já fez algum exame aqui?", "Puxar Internado", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[]{"Sim", "Não"}, "Sim");
            if (resp == JOptionPane.YES_OPTION) {
                int respBusca = JOptionPane.showOptionDialog(this, "Buscar por CNS ou selecionar arquivo?", "Paciente Existente", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[]{"CNS", "Arquivo"}, "CNS");
                File fichaLocal = null;
                String nomePacienteFallback = null;
                if (respBusca == JOptionPane.YES_OPTION) {
                    String cns = JOptionPane.showInputDialog(this, "Digite o CNS do paciente:");
                    if (cns != null && !cns.trim().isEmpty()) {
                        setStatus("Buscando paciente por CNS...");
                        fichaLocal = buscarArquivoPorCNS(cns.trim());
                        if (fichaLocal == null) {
                            JOptionPane.showMessageDialog(this, "Nenhum paciente encontrado com este CNS.", "Erro", JOptionPane.ERROR_MESSAGE);
                            setStatus("Pronto.");
                            return;
                        }
                        setStatus("Paciente encontrado: " + fichaLocal.getName());
                    } else return;
                } else {
                    JFileChooser chooser = new JFileChooser(new File(rootPath));
                    chooser.setDialogTitle("Selecione o Laudo Anterior do Paciente");
                    chooser.setFileFilter(new FileNameExtensionFilter("Documentos Word (*.docx, *.doc)", "docx", "doc"));
                    if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                        fichaLocal = chooser.getSelectedFile();
                    } else return;
                }

                if (fichaLocal != null) {
                    Map<String, String> dadosFicha = extrairDados(fichaLocal);
                    nomePacienteFallback = dadosFicha.get("{{NOME}}");
                    if (nomePacienteFallback == null || nomePacienteFallback.equals("Nome não identificado")) {
                        nomePacienteFallback = JOptionPane.showInputDialog(this, "Não foi possível extrair o nome. Digite o nome completo:");
                        if (nomePacienteFallback == null || nomePacienteFallback.trim().isEmpty()) return;
                    }
                    
                    int respTipo = JOptionPane.showOptionDialog(this, "Deseja gerar Laudo para CAT ou ANGIOPLASTIA?", "Tipo de Laudo", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[]{"CAT", "ANGIO"}, "CAT");
                    String tipo = (respTipo == JOptionPane.YES_OPTION) ? "cateterismo" : "angioplastia";
                    String numExame = calcularProximoNumero();
                    String numConfirmado = JOptionPane.showInputDialog(this, "Número do novo exame:", numExame);
                    if (numConfirmado == null) return;

                    File laudo = gerarDocumento(tipo, numConfirmado, fichaLocal.getName(), nomePacienteFallback, tipo.equals("angioplastia"));
                    abrirArquivo(laudo);
                    setStatus("Laudo de evolução gerado com sucesso.");
                }
            } else if (resp == JOptionPane.NO_OPTION) {
                String nomePaciente = JOptionPane.showInputDialog(this, "Digite o nome completo do paciente:");
                if (nomePaciente != null && !nomePaciente.trim().isEmpty()) {
                     int respTipo = JOptionPane.showOptionDialog(this, "Deseja gerar Laudo para CAT ou ANGIOPLASTIA?", "Tipo de Laudo", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[]{"CAT", "ANGIO"}, "CAT");
                    String tipo = (respTipo == JOptionPane.YES_OPTION) ? "cateterismo" : "angioplastia";
                    String numExame = calcularProximoNumero();
                     String numConfirmado = JOptionPane.showInputDialog(this, "Número do novo exame:", numExame);
                    if (numConfirmado == null) return;
                    File laudo = gerarDocumentoLocal(tipo, numConfirmado, nomePaciente, new HashMap<>(), tipo.equals("angioplastia"));
                    abrirArquivo(laudo);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Ocorreu um erro: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
            setStatus("Erro.");
        }
    }

    private void fluxoInteligente() {
        int row = tablePendentes.getSelectedRow();
        if (row == -1) {
            JOptionPane.showMessageDialog(this, "Selecione um paciente na aba Pendentes!");
            return;
        }
        String nomeArquivo = (String) modelPendentes.getValueAt(row, 1);
        String nomePacienteLista = (String) modelPendentes.getValueAt(row, 0);
        try {
            setStatus("Baixando ficha...");
            File fichaTemp = new File("temp_leitura_" + System.currentTimeMillis() + ".docx");
            downloadFile(serverUrl + "/api/baixar/" + encodePath(nomeArquivo), fichaTemp);
            Map<String, String> dados = extrairDados(fichaTemp);
            fichaTemp.delete();
            String tipoProcedimento = dados.get("{{PROCEDIMENTO}}");
            String numExame = calcularProximoNumero();
            String numConfirmado = JOptionPane.showInputDialog(this, "Procedimento detectado: " + tipoProcedimento + "\nNúmero do Exame:", numExame);
            if (numConfirmado == null) return;
            if (tipoProcedimento.toUpperCase().contains("ANGIOPLASTIA")) {
                setStatus("Gerando Angioplastia Direta...");
                File laudo = gerarDocumento("angioplastia", numConfirmado, nomeArquivo, nomePacienteLista, false);
                notifyServerDone(nomeArquivo);
                atualizarTabelasPosGeracao(row);
                abrirArquivo(laudo);
                setStatus("Angioplastia Direta Concluída.");
            } else {
                setStatus("Gerando Cateterismo...");
                File laudoCat = gerarDocumento("cateterismo", numConfirmado, nomeArquivo, nomePacienteLista, false);
                notifyServerDone(nomeArquivo);
                atualizarTabelasPosGeracao(row);
                abrirArquivo(laudoCat);
                int resp = JOptionPane.showConfirmDialog(this, "Deseja gerar também o laudo de ANGIOPLASTIA (PTCA)?\n(Será o exame nº " + incrementarNumero(numConfirmado) + ")", "Evolução para Intervenção", JOptionPane.YES_NO_OPTION);
                if (resp == JOptionPane.YES_OPTION) {
                    String numAngio = incrementarNumero(numConfirmado);
                    setStatus("Gerando Angioplastia (Evolução)...");
                    File laudoAngio = gerarDocumento("angioplastia", numAngio, nomeArquivo, nomePacienteLista, true);
                    abrirArquivo(laudoAngio);
                }
                setStatus("Fluxo Cateterismo finalizado.");
            }
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Erro: " + e.getMessage());
        }
    }

    private void fluxoAngioplastiaHistorico() {
        int row = tableHistorico.getSelectedRow();
        if (row == -1) {
            JOptionPane.showMessageDialog(this, "Selecione um arquivo no Histórico!");
            return;
        }
        String nomeArquivo = (String) modelHistorico.getValueAt(row, 0);
        try {
            String numSugestao = calcularProximoNumero();
            String numConfirmado = JOptionPane.showInputDialog(this, "Número da Angioplastia:", numSugestao);
            if (numConfirmado == null) return;
            setStatus("Gerando Angioplastia do Histórico...");
            File laudo = gerarDocumento("angioplastia", numConfirmado, nomeArquivo, null, true);
            abrirArquivo(laudo);
            setStatus("Angioplastia gerada.");
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Erro: " + e.getMessage());
        }
    }

    private File gerarDocumento(String tipo, String numeroExame, String nomeArquivoFicha, String nomePacienteFallback, boolean usarSufixoPtcaNaPasta) throws Exception {
        File fichaTemp = new File("temp_ficha_" + System.currentTimeMillis() + ".docx");
        downloadFile(serverUrl + "/api/baixar/" + encodePath(nomeArquivoFicha), fichaTemp);

        // Se o arquivo já for um laudo pré-pronto gerado pelo Python
        if (nomeArquivoFicha.toUpperCase().contains("_CATETERISMO") || nomeArquivoFicha.toUpperCase().contains("_ANGIOPLASTIA")) {
            String nomePaciente = extrairNomeDoArquivo(nomeArquivoFicha);
            if (nomePaciente == null || nomePaciente.isEmpty()) {
                nomePaciente = nomePacienteFallback;
            }

            File pastaFinal = criarEstruturaDePastas(numeroExame, nomePaciente, usarSufixoPtcaNaPasta);

            Map<String, String> dados = new HashMap<>();
            dados.put("{{NUM_EXAME}}", numeroExame);
            dados.put("{{DATA_HOJE}}", LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

            try (FileInputStream fis = new FileInputStream(fichaTemp);
                 XWPFDocument doc = new XWPFDocument(fis)) {

                substituirPlaceholders(doc, dados);

                String nomeArquivoSaida = nomePaciente.replaceAll("[^a-zA-Z0-9 .\\-]", "") + 
                                         (nomeArquivoFicha.toUpperCase().contains("_ANGIOPLASTIA") && usarSufixoPtcaNaPasta ? " PTCA" : "") + ".docx";
                File arquivoFinal = new File(pastaFinal, nomeArquivoSaida);

                try (FileOutputStream fos = new FileOutputStream(arquivoFinal)) {
                    doc.write(fos);
                }

                fichaTemp.delete();
                lastGeneratedFile = arquivoFinal;
                btnAbrirUltimo.setEnabled(true);
                return arquivoFinal;
            }
        }

        // Fluxo normal para fichas de papel digitalizadas
        Map<String, String> dados = extrairDados(fichaTemp);
        String nomePaciente = dados.get("{{NOME}}");
        if (nomePaciente == null || nomePaciente.equals("Nome não identificado")) {
            nomePaciente = nomePacienteFallback;
        }

        File laudo = gerarDocumentoLocal(tipo, numeroExame, nomePaciente, dados, usarSufixoPtcaNaPasta);
        fichaTemp.delete();
        return laudo;
    }

    private String extrairNomeDoArquivo(String nomeArquivo) {
        try {
            // Formato esperado: 20260309_Nome_Paciente_TIPO.docx
            String semExtensao = nomeArquivo.substring(0, nomeArquivo.lastIndexOf('.'));
            String[] partes = semExtensao.split("_");
            if (partes.length >= 2) {
                StringBuilder nome = new StringBuilder();
                for (int i = 1; i < partes.length; i++) {
                    if (partes[i].equalsIgnoreCase("CATETERISMO") || partes[i].equalsIgnoreCase("ANGIOPLASTIA")) break;
                    if (nome.length() > 0) nome.append(" ");
                    nome.append(partes[i]);
                }
                return toPascalCase(nome.toString().trim());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "Paciente";
    }
    private File gerarDocumentoLocal(String tipo, String numeroExame, String nomePaciente, Map<String, String> dados, boolean usarSufixoPtcaNaPasta) throws Exception {
        File modeloTemp = new File("temp_modelo_" + tipo + ".docx");
        downloadFile(serverUrl + "/api/modelos/" + (tipo.equals("cateterismo") ? "Laudo de cateterismo.docx" : "Laudo de Angioplastia.docx"), modeloTemp);

        dados.put("{{NUM_EXAME}}", numeroExame);
        dados.put("{{DATA_HOJE}}", LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        dados.put("{{NOME}}", toPascalCase(nomePaciente));
        dados.put("{{PROCEDENCIA}}", toPascalCase(dados.get("{{PROCEDENCIA}}")));

        File pastaFinal = criarEstruturaDePastas(numeroExame, nomePaciente, usarSufixoPtcaNaPasta);
        FileInputStream fis = new FileInputStream(modeloTemp);
        XWPFDocument doc = new XWPFDocument(fis);
        substituirPlaceholders(doc, dados);
        fis.close();
        String nomeArquivoSaida = nomePaciente.replaceAll("[^a-zA-Z0-9 .\\-]", "") + (tipo.equals("angioplastia") && usarSufixoPtcaNaPasta ? " PTCA" : "") + ".docx";
        File arquivoFinal = new File(pastaFinal, nomeArquivoSaida);
        FileOutputStream fos = new FileOutputStream(arquivoFinal);
        doc.write(fos);
        fos.close();
        doc.close();
        modeloTemp.delete();
        lastGeneratedFile = arquivoFinal;
        btnAbrirUltimo.setEnabled(true);
        return arquivoFinal;
    }
    
    private void setupPendentesRightClick() {
        JPopupMenu popupMenu = new JPopupMenu();
        JMenuItem deleteItem = new JMenuItem("Excluir");
        deleteItem.addActionListener(e -> {
            int selectedRow = tablePendentes.getSelectedRow();
            if (selectedRow != -1) {
                excluirPendente(selectedRow);
            }
        });
        popupMenu.add(deleteItem);

        tablePendentes.addMouseListener(new MouseAdapter() {
            public void mousePressed(MouseEvent e) {
                if (SwingUtilities.isRightMouseButton(e)) {
                    int row = tablePendentes.rowAtPoint(e.getPoint());
                    if (row >= 0 && row < tablePendentes.getRowCount()) {
                        tablePendentes.setRowSelectionInterval(row, row);
                        popupMenu.show(e.getComponent(), e.getX(), e.getY());
                    }
                }
            }
        });
    }

    private void excluirPendente(int row) {
        String arquivo = (String) modelPendentes.getValueAt(row, 1);
        int confirm = JOptionPane.showConfirmDialog(this, "Tem certeza que deseja excluir o pendente?\n" + arquivo, "Excluir Pendente", JOptionPane.YES_NO_OPTION);
        if (confirm != JOptionPane.YES_OPTION) {
            return;
        }

        new Thread(() -> {
            try {
                String urlStr = serverUrl + "/api/excluir/" + encodePath(arquivo);
                HttpURLConnection conn = (HttpURLConnection) new URL(urlStr).openConnection();
                conn.setRequestMethod("DELETE");
                conn.setRequestProperty("User-Agent", "Mozilla/5.0");

                if (conn.getResponseCode() == 200) {
                    SwingUtilities.invokeLater(() -> {
                        modelPendentes.removeRow(row);
                        JOptionPane.showMessageDialog(this, "Pendente excluído com sucesso!", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
                        setStatus("Pendente removido.");
                    });
                } else {
                    throw new IOException("Erro no servidor: " + conn.getResponseCode());
                }
            } catch (Exception e) {
                e.printStackTrace();
                SwingUtilities.invokeLater(() -> {
                    JOptionPane.showMessageDialog(this, "Erro ao excluir pendente: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
                    setStatus("Falha ao remover pendente.");
                });
            }
        }).start();
    }
    
    private void excluirRegistrosHistorico() {
        int[] selectedRows = tableHistorico.getSelectedRows();
        if (selectedRows.length == 0) {
            JOptionPane.showMessageDialog(this, "Selecione um ou mais arquivos no Histórico para excluir.");
            return;
        }

        int confirm = JOptionPane.showConfirmDialog(this, "Tem certeza que deseja excluir " + selectedRows.length + " registro(s)?", "Excluir", JOptionPane.YES_NO_OPTION);
        if (confirm != JOptionPane.YES_OPTION) {
            return;
        }

        new Thread(() -> {
            List<String> arquivosParaExcluir = new ArrayList<>();
            for (int row : selectedRows) {
                arquivosParaExcluir.add((String) modelHistorico.getValueAt(row, 0));
            }

            int sucesso = 0;
            int falha = 0;
            List<String> falhas = new ArrayList<>();

            for (String arquivo : arquivosParaExcluir) {
                try {
                    String urlStr = serverUrl + "/api/excluir/" + encodePath(arquivo);
                    HttpURLConnection conn = (HttpURLConnection) new URL(urlStr).openConnection();
                    conn.setRequestMethod("DELETE");
                    conn.setRequestProperty("User-Agent", "Mozilla/5.0");

                    if (conn.getResponseCode() == 200) {
                        sucesso++;
                    } else {
                        falha++;
                        falhas.add(arquivo);
                    }
                } catch (Exception e) {
                    falha++;
                    falhas.add(arquivo);
                    e.printStackTrace();
                }
            }

            int finalSucesso = sucesso;
            int finalFalha = falha;
            SwingUtilities.invokeLater(() -> {
                refreshListHistorico();
                String msg = finalSucesso + " registro(s) excluído(s) com sucesso.";
                if (finalFalha > 0) {
                    msg += "\n" + finalFalha + " falharam.";
                    System.err.println("Falha ao excluir: " + String.join(", ", falhas));
                }
                JOptionPane.showMessageDialog(this, msg, "Resultado da Exclusão", JOptionPane.INFORMATION_MESSAGE);
                setStatus(finalSucesso + " excluídos, " + finalFalha + " falhas.");
            });
        }).start();
    }
    
    private void refreshListPendentes() {
        new Thread(() -> {
            try {
                String json = httpGet(serverUrl + "/api/pendentes");
                JSONArray arr = new JSONArray(json);
                SwingUtilities.invokeLater(() -> {
                    modelPendentes.setRowCount(0);
                    for (int i = 0; i < arr.length(); i++) {
                        JSONObject obj = arr.getJSONObject(i);
                        modelPendentes.addRow(new Object[]{
                            obj.getString("nome"), 
                            obj.getString("arquivo"),
                            obj.getString("procedencia"),
                            obj.getString("tipo_procedimento")
                        });
                    }
                    setStatus(arr.length() + " pendentes carregados.");
                });
            } catch (Exception e) {
                e.printStackTrace();
                SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(this, "Erro ao atualizar pendentes: " + e.getMessage()));
            }
        }).start();
    }

    private void refreshListHistorico() {
        new Thread(() -> {
            try {
                String json = httpGet(serverUrl + "/api/historico");
                JSONArray arr = new JSONArray(json);
                SwingUtilities.invokeLater(() -> {
                    modelHistorico.setRowCount(0);
                    for (int i = 0; i < arr.length(); i++) {
                        modelHistorico.addRow(new Object[]{arr.getJSONObject(i).getString("arquivo")});
                    }
                    setStatus(arr.length() + " registros de histórico carregados.");
                });
            } catch (Exception e) {
                e.printStackTrace();
            }
        }).start();
    }

    private void refreshListInternacao() {
        new Thread(() -> {
            try {
                String json = httpGet(serverUrl + "/api/internacao/listar");
                JSONArray arr = new JSONArray(json);
                SwingUtilities.invokeLater(() -> {
                    modelInternacao.setRowCount(0);
                    for (int i = 0; i < arr.length(); i++) {
                        JSONObject obj = arr.getJSONObject(i);
                        modelInternacao.addRow(new Object[]{
                            obj.getString("nome"),
                            obj.getString("data"),
                            obj.getString("arquivo"),
                            obj.optString("procedencia", "Não informada") // (Ponto 2) Captura a procedência
                        });
                    }
                    setStatus(arr.length() + " pacientes para internação carregados.");
                });
            } catch (Exception e) {
                e.printStackTrace();
                SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(this, "Erro ao carregar lista de internação: " + e.getMessage()));
            }
        }).start();
    }

    private void abrirDialogoGeracaoDocs() {
        int selectedRow = tableInternacao.getSelectedRow();
        if (selectedRow == -1) {
            JOptionPane.showMessageDialog(this, "Selecione um paciente na lista de internação.", "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }
    
        String nomePaciente = (String) modelInternacao.getValueAt(selectedRow, 0);
        String procedencia = (String) modelInternacao.getValueAt(selectedRow, 3); // (Ponto 2) Pega a procedência
    
        // --- Main Panel ---
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
    
        // --- Standard Checkboxes ---
        JCheckBox chkReceita = new JCheckBox("Receita AAS + Clopidogrel", true);
        JCheckBox chkInternacao = new JCheckBox("Solicitação de Internação", true);
        JCheckBox chkSolAngio = new JCheckBox("Solicitação de Angioplastia (SUS)", true);
        JCheckBox chkJustAngio = new JCheckBox("Justificativa para Angioplastia", true);
        JCheckBox chkEvolucao = new JCheckBox("Evolução Tasy", true);
        JCheckBox chkTransporte = new JCheckBox("Ficha de Transporte (em desenvolvimento)");
        chkTransporte.setEnabled(false);
    
        // --- General Input Fields (for Solicitação de Internação, etc.) ---
        JPanel generalPanel = new JPanel(new GridLayout(0, 2, 5, 5));
        generalPanel.setBorder(BorderFactory.createTitledBorder("Dados Gerais (Solicitação de Internação)"));
        generalPanel.add(new JLabel("Artérias Tratadas:"));
        JTextField arteriasField = new JTextField(20);
        generalPanel.add(arteriasField);
        generalPanel.add(new JLabel("Quantidade de Stents:"));
        JTextField stentsField = new JTextField(5);
        generalPanel.add(stentsField);
    
        // --- (Ponto 10) Panel for Justificativa Options (conditionally visible) ---
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
    
        // --- "Sugestão de Conduta" Section ---
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
    
        // --- Assembling the final panel ---
        panel.add(chkReceita);
        panel.add(chkInternacao);
        panel.add(chkSolAngio);
        panel.add(chkJustAngio);
        panel.add(chkEvolucao);
        panel.add(chkTransporte);
        panel.add(Box.createVerticalStrut(10));
        panel.add(generalPanel);
        panel.add(justificativaOptionsPanel);
        panel.add(Box.createVerticalStrut(10));
        panel.add(chkConduta);
        panel.add(condutaOptionsPanel);
    
        // (Ponto 7) Wrap in a scroll pane and set preferred size for a larger dialog
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
            payloadData.put("arterias", arteriasField.getText());
            payloadData.put("stents", stentsField.getText());
            payloadData.put("procedencia", procedencia); // (Ponto 2) Adiciona a procedência ao payload
    
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
        if (modelos.isEmpty()) return;

        new Thread(() -> {
            try {
                URL url = new URL(serverUrl + "/api/internacao/gerar");
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("POST");
                conn.setRequestProperty("Content-Type", "application/json");
                conn.setRequestProperty("Accept", "application/json");
                conn.setRequestProperty("User-Agent", "Mozilla/5.0");
                conn.setDoOutput(true);

                JSONObject payload = new JSONObject();
                payload.put("nome", nome);
                payload.put("modelos", new JSONArray(modelos));

                for (Map.Entry<String, String> entry : data.entrySet()) {
                    payload.put(entry.getKey(), entry.getValue());
                }

                try (OutputStream os = conn.getOutputStream()) {
                    os.write(payload.toString().getBytes(StandardCharsets.UTF_8));
                }

                if (conn.getResponseCode() != 201) {
                    String errorResponse = "";
                    try (BufferedReader br = new BufferedReader(new InputStreamReader(conn.getErrorStream(), StandardCharsets.UTF_8))) {
                        errorResponse = br.lines().collect(Collectors.joining("\n"));
                    } catch (Exception e) {
                        // Ignora se não conseguir ler o erro
                    }
                    throw new IOException("Erro no servidor ao gerar documentos: " + conn.getResponseCode() + " - " + errorResponse);
                }

                try (BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
                    JSONObject response = new JSONObject(br.lines().collect(Collectors.joining("\n")));
                    JSONArray arquivosGerados = response.getJSONArray("arquivos_gerados");

                    for (int i = 0; i < arquivosGerados.length(); i++) {
                        String filename = arquivosGerados.getString(i);
                        File tempFile = File.createTempFile("internacao_", ".docx");
                        tempFile.deleteOnExit();

                        downloadFile(serverUrl + "/api/baixar/" + encodePath(filename), tempFile);
                        abrirArquivo(tempFile);
                    }
                    SwingUtilities.invokeLater(() -> setStatus(arquivosGerados.length() + " documentos abertos para impressão."));
                }

            } catch (Exception e) {
                e.printStackTrace();
                SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(this, "Falha ao gerar/imprimir documentos: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE));
            }
        }).start();
    }
    
    private void refreshAll() {
        refreshListPendentes();
        refreshListHistorico();
        refreshListInternacao();
    }
    
    private String httpGet(String urlStr) throws IOException {
        URL url = new URL(urlStr);
        HttpURLConnection http = (HttpURLConnection) url.openConnection();
        http.setRequestProperty("User-Agent", "Mozilla/5.0");
        http.setRequestProperty("Accept", "application/json");
        http.setConnectTimeout(5000);

        if (http.getResponseCode() >= 400) {
            throw new IOException("Erro Server: " + http.getResponseCode());
        }

        try (Scanner sc = new Scanner(http.getInputStream(), "UTF-8")) {
            sc.useDelimiter("\\A");
            return sc.hasNext() ? sc.next() : "";
        }
    }
    
    private void downloadFile(String urlStr, File dest) throws IOException {
        URL url = new URL(urlStr);
        HttpURLConnection http = (HttpURLConnection) url.openConnection();
        http.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
        http.setConnectTimeout(10000);

        int responseCode = http.getResponseCode();
        if (responseCode != 200) {
            throw new IOException("Erro Download (" + responseCode + "): " + urlStr);
        }

        InputStream in = http.getInputStream();
        Files.copy(in, dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
        in.close();
    }
    
    private String encodePath(String filename) {
        try {
            if (filename == null) return "";
            return URLEncoder.encode(filename, StandardCharsets.UTF_8.toString()).replace("+", "%20");
        } catch (UnsupportedEncodingException e) {
            return filename.replace(" ", "%20"); // Fallback
        }
    }
    
    private Map<String, String> extrairDados(File arquivo) throws IOException {
        Map<String, String> dados = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(arquivo)) {
            XWPFDocument doc = new XWPFDocument(fis);
            String textoCompleto = extrairTextoCompleto(doc);

            dados.put("{{NOME}}", regexExtract(textoCompleto, "(?:NOME|PACIENTE|NM)\\s*[:\\s_.]+\\s*([A-ZÀ-Ú\\s]{5,})(?:\n|\\s{2,}|DATA|NASC|END)"));
            dados.put("{{CNS}}", regexExtract(textoCompleto, "(?:CNS|MATRÍCULA|CARTÃO SUS|CNS:):\\s*([\\d\\.\\-\\s]+)"));
            dados.put("{{NASC}}", regexExtract(textoCompleto, "NASCIMENTO:\\s*([\\d]{2}/[\\d]{2}/[\\d]{4})"));
            dados.put("{{PROCEDENCIA}}", regexExtract(textoCompleto, "(?:UNIDADE DE ORIGEM|PROCEDENCIA)[:\\s_]+(.*?)(?:\\n|M.DICO|CONV.NIO|LEITO)"));
            
            if (textoCompleto.toUpperCase().contains("ANGIOPLASTIA") || textoCompleto.toUpperCase().contains("PTCA")) {
                dados.put("{{PROCEDIMENTO}}", "ANGIOPLASTIA");
            } else {
                dados.put("{{PROCEDIMENTO}}", "CATETERISMO");
            }
        }
        return dados;
    }
    
    private void abrirArquivo(File arquivo) {
        if (!Desktop.isDesktopSupported() || arquivo == null) return;

        new Thread(() -> {
            int attempts = 0;
            while (attempts < 5) {
                try {
                    Thread.sleep(1000 + (attempts * 500));
                    if (arquivo.exists()) {
                        Desktop.getDesktop().open(arquivo);
                        return;
                    }
                } catch (IOException e) {
                    System.err.println("Tentativa " + (attempts + 1) + " falhou: " + e.getMessage());
                } catch (InterruptedException ie) {
                    Thread.currentThread().interrupt();
                    return;
                }
                attempts++;
            }
            SwingUtilities.invokeLater(() -> {
                setStatus("Falha ao abrir o arquivo automaticamente.");
                JOptionPane.showMessageDialog(this,
                        "Não foi possível abrir o arquivo automaticamente.\nUse o botão 'Abrir Último Laudo' para tentar novamente.",
                        "Falha ao Abrir",
                        JOptionPane.WARNING_MESSAGE);
            });
        }).start();
    }
    
    private void loadConfig() {
        try {
            File f = new File(CONFIG_FILE);
            if (f.exists()) {
                JSONObject j = new JSONObject(new String(Files.readAllBytes(f.toPath())));
                serverUrl = j.optString("url", serverUrl);
                rootPath = j.optString("path", rootPath);
                lastVersionSeen = j.optString("lastVersion", "");
            }
        } catch (Exception ignored) {}
    }

    private void saveConfig() {
        try {
            JSONObject json = new JSONObject();
            json.put("url", serverUrl);
            json.put("path", rootPath);
            json.put("lastVersion", lastVersionSeen);
            Files.write(Paths.get(CONFIG_FILE), json.toString(4).getBytes());
        } catch (Exception ignored) {}
    }
    
    private void showConfigDialog() {
        JTextField urlField = new JTextField(serverUrl);
        JTextField pathField = new JTextField(rootPath);
        Object[] message = {
                "URL do Servidor:", urlField,
                "Pasta Raiz (Laudos):", pathField
        };

        int option = JOptionPane.showConfirmDialog(this, message, "Configurações", JOptionPane.OK_CANCEL_OPTION);
        if (option == JOptionPane.OK_OPTION) {
            serverUrl = urlField.getText().trim().replaceAll("/$", "");
            rootPath = pathField.getText().trim();
            infoLabel.setText("Conectado a: " + serverUrl);
            saveConfig();
            refreshAll();
        }
    }
    
    private String toPascalCase(String input) {
        if (input == null || input.isEmpty()) return "";
        // (Ponto 1) Lista de preposições a serem mantidas em minúsculo.
        List<String> prepositions = Arrays.asList("de", "da", "do", "di", "du");
        String[] words = input.toLowerCase().split("\\s+");
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < words.length; i++) {
            String w = words[i];
            if (w.length() > 0) {
                // Capitaliza a primeira palavra ou se não for uma preposição.
                if (i == 0 || !prepositions.contains(w)) {
                    sb.append(Character.toUpperCase(w.charAt(0))).append(w.substring(1));
                } else {
                    sb.append(w);
                }
                sb.append(" ");
            }
        }
        return sb.toString().trim();
    }
    
    private String incrementarNumero(String atual) {
        try {
            int val = Integer.parseInt(atual.replace(".", "")) + 1;
            String s = String.valueOf(val);
            return s.length() > 3 ? s.substring(0, s.length() - 3) + "." + s.substring(s.length() - 3) : s;
        } catch (Exception e) {
            return atual + " (Verificar)";
        }
    }
    
    private String calcularProximoNumero() {
        File pastaMes = getPastaMes();
        List<File> locais = Arrays.asList(pastaMes, new File(pastaMes, "PACIENTE"));
        int maior = 0;
        Pattern p = Pattern.compile("(\\d{2})\\.(\\d{3})");
        for (File dir : locais) {
            if (dir.exists() && dir.isDirectory()) {
                File[] files = dir.listFiles();
                if (files != null) {
                    for (File f : files) {
                        if (f.isDirectory()) {
                            Matcher m = p.matcher(f.getName());
                            if (m.find()) {
                                try {
                                    int num = Integer.parseInt(m.group(1) + m.group(2));
                                    if (num > maior) maior = num;
                                } catch (Exception ignored) {}
                            }
                        }
                    }
                }
            }
        }
        if (maior == 0) return "42.001";
        int prox = maior + 1;
        String s = String.valueOf(prox);
        return s.substring(0, 2) + "." + s.substring(2);
    }
    
    private File getPastaMes() {
        LocalDate hoje = LocalDate.now();
        String[] meses = {"", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"};
        return new File(rootPath, String.format("%02d %s", hoje.getMonthValue(), meses[hoje.getMonthValue()]));
    }
    
    private void setStatus(String msg) {
        statusLabel.setText(msg);
    }
    
    private void notifyServerDone(String filename) {
        new Thread(() -> {
            try {
                String encoded = encodePath(filename);
                URL url = new URL(serverUrl + "/api/concluir/" + encoded);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("POST");
                conn.setRequestProperty("User-Agent", "Mozilla/5.0");
                conn.getResponseCode();
            } catch (Exception ignored) {}
        }).start();
    }
    
    private void atualizarTabelasPosGeracao(int row) {
        SwingUtilities.invokeLater(() -> {
            if (row < modelPendentes.getRowCount()) modelPendentes.removeRow(row);
            refreshListHistorico();
        });
    }
    
    private File criarEstruturaDePastas(String numeroExame, String nomePaciente, boolean usarSufixoPtcaNaPasta) {
        File pastaMes = getPastaMes();
        File pastaPacientes = new File(pastaMes, "PACIENTE");
        if (!pastaPacientes.exists()) pastaPacientes.mkdirs();

        String nomePacienteLimpo = nomePaciente.replaceAll("[^a-zA-Z0-9 .\\-]", "");
        String sufixoPasta = usarSufixoPtcaNaPasta ? " PTCA" : "";
        String nomePasta = numeroExame + " " + nomePacienteLimpo + sufixoPasta;
        File pastaFinal = new File(pastaPacientes, nomePasta);
        pastaFinal.mkdirs();
        return pastaFinal;
    }
    
    private void substituirPlaceholders(XWPFDocument doc, Map<String, String> dados) {
        // (Pontos 3, 6, 8, 9) Substituição que preserva a formatação.
        for (XWPFParagraph p : doc.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                String text = r.getText(0);
                if (text != null) {
                    for (Map.Entry<String, String> entry : dados.entrySet()) {
                        if (text.contains(entry.getKey())) {
                            text = text.replace(entry.getKey(), entry.getValue() != null ? entry.getValue() : "");
                        }
                    }
                    r.setText(text, 0);
                }
            }
        }
        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null) {
                                for (Map.Entry<String, String> entry : dados.entrySet()) {
                                    if (text.contains(entry.getKey())) {
                                        text = text.replace(entry.getKey(), entry.getValue() != null ? entry.getValue() : "");
                                    }
                                }
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }

    private void gerarRelatorioDiario() {
        setStatus("Gerando relatório...");
        new Thread(() -> {
            File pastaMes = getPastaMes();
            File pastaPacientes = new File(pastaMes, "PACIENTE");
            List<Object[]> linhas = new ArrayList<>();
            LocalDate hoje = LocalDate.now();

            if (pastaPacientes.exists()) {
                File[] pastasExames = pastaPacientes.listFiles(File::isDirectory);
                if (pastasExames != null) {
                    Arrays.sort(pastasExames, Comparator.comparing(File::getName));
                    for (File pastaExame : pastasExames) {
                        if (foiModificadoHoje(pastaExame, hoje)) {
                            analizarPastaExame(pastaExame, linhas);
                        }
                    }
                }
            }
            SwingUtilities.invokeLater(() -> {
                modelRelatorio.setRowCount(0);
                for (Object[] row : linhas) modelRelatorio.addRow(row);
                setStatus("Relatório atualizado.");
            });
        }).start();
    }

    private void analizarPastaExame(File pasta, List<Object[]> linhas) {
        File[] docs = pasta.listFiles((d, name) -> name.endsWith(".docx") && !name.startsWith("~$"));
        if (docs == null) return;
        
        for (File doc : docs) {
            try {
                Map<String, String> dados = extrairDados(doc);
                linhas.add(new Object[]{
                    dados.get("{{NOME}}"), 
                    dados.get("{{PROCEDENCIA}}"), 
                    dados.get("{{PROCEDIMENTO}}"), 
                    "-"
                });
            } catch (Exception e) {
                // Ignora
            }
        }
    }

    private boolean foiModificadoHoje(File f, LocalDate hoje) {
        try {
            BasicFileAttributes attr = Files.readAttributes(f.toPath(), BasicFileAttributes.class);
            LocalDate dm = attr.lastModifiedTime().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            return dm.equals(hoje);
        } catch(Exception e) { return false; }
    }

    private File buscarArquivoPorCNS(String cns) throws IOException {
        final File[] encontrado = {null};
        String cnsNumerico = cns.replaceAll("[^\\d]", "");

        Files.walk(Paths.get(rootPath))
                .filter(path -> path.toString().toLowerCase().endsWith(".docx") && !path.getFileName().toString().startsWith("~$"))
                .parallel()
                .forEach(path -> {
                    if (encontrado[0] == null) {
                        try (FileInputStream fis = new FileInputStream(path.toFile())) {
                            XWPFDocument doc = new XWPFDocument(fis);
                            String texto = extrairTextoCompleto(doc);
                            doc.close();
                            String cnsNoDoc = regexExtract(texto, "(?:CNS|MATRÍCULA|CARTÃO SUS|CNS:):\\s*([\\d\\.\\-\\s]+)").replaceAll("[^\\d]", "");
                            if (cnsNumerico.equals(cnsNoDoc)) {
                                encontrado[0] = path.toFile();
                            }
                        } catch (Exception e) {
                            // Ignora
                        }
                    }
                });

        return encontrado[0];
    }
    
    private String extrairTextoCompleto(XWPFDocument doc) {
        StringBuilder text = new StringBuilder();
        for (XWPFParagraph p : doc.getParagraphs()) {
            text.append(p.getText()).append("\n");
        }
        for (XWPFTable t : doc.getTables()) {
            for (XWPFTableRow r : t.getRows()) {
                for (XWPFTableCell c : r.getTableCells()) {
                    text.append(c.getText()).append("\t");
                }
                text.append("\n");
            }
        }
        return text.toString();
    }
    
    private String regexExtract(String text, String patternStr) {
        Pattern p = Pattern.compile(patternStr, Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
        Matcher m = p.matcher(text);
        return m.find() ? m.group(1).trim().replace("\n", " ") : "";
    }
}
