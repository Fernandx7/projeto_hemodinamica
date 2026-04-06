import javax.swing.*;
import javax.swing.border.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.MaskFormatter;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONObject;
import org.json.JSONArray;

public class AppRecepcao extends JFrame {

    // --- COMPONENTES UI ---
    private JComboBox<String> comboMedico;
    private JRadioButton rbCateterismo, rbAngioplastia;
    private JTextField txtNome, txtLogradouro, txtBairro, txtCidade, txtUf;
    private JFormattedTextField txtNasc, txtRg, txtCpf, txtCep;
    private JTextField txtMae, txtPai, txtMatricula, txtChave;
    private JPanel panelTelefones;
    private List<JPanel> listaPaineisTelefone = new ArrayList<>();
    private JPanel glassPane;
    private String rascunhoAberto = null;
    private boolean isDraftMode = false;

    // Controle de Abas
    private JRadioButton rbResidencia, rbAmbulancia;
    private JPanel panelDinamico;
    private CardLayout cardLayout;

    // Campos Eletivo
    private JTextField txtPeso, txtAltura;
    private JCheckBox chkIodo, chkAnti, chkDiab, chkCiru;
    private JRadioButton rbPreparoSim, rbPreparoNao;
    private JTextField txtAntiNome, txtDiabNome, txtCiruNome;
    private JCheckBox chkSusp48, chkSusp24;

    // Campos Ambulância
    private JTextField txtHospitalOrigem;

    // CONFIG
    private static final String CONFIG_FILE = "config_recepcao.json";
    private String serverUrl = "https://hemodinamica.souzadev.software";
    private String outputDirCat = "arquivos_cateterismo";
    private String outputDirAngio = "arquivos_angioplastia";

    private final String[] LISTA_MEDICOS = {
            "DR. JOÃO FRIGHETTO", "DR. RAFAEL HENUD", "DR. ARTUR CHIMELLI",
            "DR. IGOR SANTOS", "DR. JOAO CARLOS", "DR. MARCIO MACRI"
    };

    public AppRecepcao() {
        this(false, null);
    }

    public AppRecepcao(boolean isDraftMode, JSONObject dadosRascunho) {
        super("SISTEMA DE RECEPÇÃO - HEMODINÂMICA v2.3");
        this.isDraftMode = isDraftMode;
        
        loadConfig();

        try {
            UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
        } catch (Exception ignored) {}

        setupUI();
        setSize(980, 780);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);

        if (dadosRascunho != null) {
            preencherComRascunho(dadosRascunho);
        }
    }

    private void setupUI() {
        JMenuBar menuBar = new JMenuBar();
        JMenu menuArq = new JMenu("Arquivo");
        JMenu menuFichasPre = new JMenu("Fichas Pré-Prontas");

        JMenuItem itemNovo = new JMenuItem("Novo Atendimento (Ficha Completa)");
        itemNovo.addActionListener(e -> new AppRecepcao(false, null).setVisible(true));

        JMenuItem itemAdicionarRascunho = new JMenuItem("Adicionar Ficha (Rascunho)");
        itemAdicionarRascunho.addActionListener(e -> new AppRecepcao(true, null).setVisible(true));

        JMenuItem itemAbrirRascunho = new JMenuItem("Abrir Ficha Pré-Pronta");
        itemAbrirRascunho.addActionListener(e -> abrirDialogoRascunhos());

        JMenuItem itemConfig = new JMenuItem("Configurações");
        itemConfig.addActionListener(e -> abrirConfiguracoes());

        menuArq.add(itemNovo);
        menuArq.addSeparator();
        menuArq.add(itemConfig);

        menuFichasPre.add(itemAdicionarRascunho);
        menuFichasPre.add(itemAbrirRascunho);

        menuBar.add(menuArq);
        menuBar.add(menuFichasPre);
        setJMenuBar(menuBar);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JScrollPane scroll = new JScrollPane(mainPanel);
        scroll.getVerticalScrollBar().setUnitIncrement(16);
        add(scroll);

        setupFormulario(mainPanel);

        JPanel pBotoes = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 10));
        
        if (isDraftMode) {
            JButton btnSalvarRascunho = new JButton(rascunhoAberto != null ? "Salvar Alterações no Rascunho" : "Salvar Rascunho");
            btnSalvarRascunho.setFont(new Font("Tahoma", Font.BOLD, 14));
            btnSalvarRascunho.setPreferredSize(new Dimension(240, 50));
            btnSalvarRascunho.addActionListener(e -> salvarRascunho());
            pBotoes.add(btnSalvarRascunho);

            if (rascunhoAberto != null) {
                JButton btnGerar = new JButton("Gravar Ficha (Finalizar)");
                btnGerar.setFont(new Font("Tahoma", Font.BOLD, 14));
                btnGerar.setPreferredSize(new Dimension(220, 50));
                btnGerar.setBackground(new Color(40, 167, 69));
                btnGerar.setForeground(Color.WHITE);
                btnGerar.addActionListener(e -> processarFicha());
                pBotoes.add(btnGerar);
            }
        } else {
            JButton btnGerar = new JButton("Gravar Ficha Completa");
            btnGerar.setFont(new Font("Tahoma", Font.BOLD, 14));
            btnGerar.setPreferredSize(new Dimension(220, 50));
            btnGerar.addActionListener(e -> processarFicha());
            pBotoes.add(btnGerar);

            JButton btnInternado = new JButton("Lançar Internado (Rápido)");
            btnInternado.setFont(new Font("Tahoma", Font.BOLD, 14));
            btnInternado.setPreferredSize(new Dimension(220, 50));
            btnInternado.addActionListener(e -> abrirDialogoInternado());
            pBotoes.add(btnInternado);
        }

        JButton btnLimpar = new JButton("Limpar Tela");
        btnLimpar.setPreferredSize(new Dimension(150, 50));
        btnLimpar.addActionListener(e -> limparCampos(mainPanel));
        pBotoes.add(btnLimpar);

        mainPanel.add(Box.createVerticalStrut(10));
        mainPanel.add(pBotoes);
        mainPanel.add(Box.createVerticalStrut(10));

        setupGlassPane();
    }

    private void setupFormulario(JPanel mainPanel) {
        JPanel pHeader = criarPainelRustico("DADOS DO PROCEDIMENTO");
        pHeader.setLayout(new GridLayout(2, 2, 5, 5));
        JPanel pProc = new JPanel(new FlowLayout(FlowLayout.LEFT));
        ButtonGroup bgProc = new ButtonGroup();
        rbCateterismo = new JRadioButton("CATETERISMO", true);
        rbAngioplastia = new JRadioButton("ANGIOPLASTIA");
        rbAngioplastia.addActionListener(e -> verificarImportacaoArquivo());
        bgProc.add(rbCateterismo); bgProc.add(rbAngioplastia);
        pProc.add(rbCateterismo); pProc.add(rbAngioplastia);
        pHeader.add(new JLabel("TIPO DE EXAME:"));
        pHeader.add(pProc);
        pHeader.add(new JLabel("MÉDICO EXECUTANTE:"));
        comboMedico = new JComboBox<>(LISTA_MEDICOS);
        pHeader.add(comboMedico);
        mainPanel.add(pHeader);

        JPanel pIdent = criarPainelRustico("IDENTIFICAÇÃO DO PACIENTE");
        pIdent.setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.gridx=0; gbc.gridy=0; pIdent.add(new JLabel("NOME COMPLETO:"), gbc);
        gbc.gridx=1; gbc.gridy=0; gbc.weightx=1.0; gbc.fill = GridBagConstraints.HORIZONTAL;
        txtNome = new JTextField(); pIdent.add(txtNome, gbc);
        gbc.gridx=2; gbc.gridy=0; gbc.weightx=0; gbc.fill = GridBagConstraints.NONE;
        pIdent.add(new JLabel("DATA NASC:"), gbc);
        gbc.gridx=3; gbc.gridy=0;
        try {
            MaskFormatter DataNasc = new MaskFormatter("##/##/####");
            DataNasc.setPlaceholderCharacter('_');
            txtNasc = new JFormattedTextField(DataNasc);
            txtNasc.setColumns(6);
        } catch(Exception e) { txtNasc = new JFormattedTextField(); }
        pIdent.add(txtNasc, gbc);
        gbc.gridx=0; gbc.gridy=1; pIdent.add(new JLabel("RG:"), gbc);
        gbc.gridx=1; gbc.gridy=1; gbc.fill=GridBagConstraints.NONE;
        try {
            MaskFormatter MaskRg = new MaskFormatter("##.###.###-#");
            MaskRg.setPlaceholderCharacter('_');
            txtRg = new JFormattedTextField(MaskRg);
            txtRg.setColumns(7);
        } catch(Exception e) { txtRg = new JFormattedTextField(); }
        pIdent.add(txtRg, gbc);
        gbc.gridx=2; gbc.gridy=1; pIdent.add(new JLabel("CPF:"), gbc);
        gbc.gridx=3; gbc.gridy=1;
        try {
            MaskFormatter MaskCpf = new MaskFormatter("###.###.###-##");
            MaskCpf.setPlaceholderCharacter('_');
            txtCpf = new JFormattedTextField(MaskCpf);
            txtCpf.setColumns(8);
        } catch(Exception e) { txtCpf = new JFormattedTextField(); }
        pIdent.add(txtCpf, gbc);
        gbc.gridx=0; gbc.gridy=2; pIdent.add(new JLabel("NOME DA MÃE:"), gbc);
        gbc.gridx=1; gbc.gridy=2; gbc.gridwidth=3; gbc.fill=GridBagConstraints.HORIZONTAL;
        txtMae = new JTextField(); pIdent.add(txtMae, gbc);
        gbc.gridwidth=1;
        gbc.gridx=0; gbc.gridy=3; pIdent.add(new JLabel("NOME DO PAI:"), gbc);
        gbc.gridx=1; gbc.gridy=3; gbc.gridwidth=3; gbc.fill=GridBagConstraints.HORIZONTAL;
        txtPai = new JTextField(); pIdent.add(txtPai, gbc);
        gbc.gridwidth=1;
        mainPanel.add(pIdent);

        JPanel pEnd = criarPainelRustico("ENDEREÇO E CONTATO");
        pEnd.setLayout(new GridBagLayout());
        gbc.gridx=0; gbc.gridy=0; gbc.weightx=0; gbc.fill=GridBagConstraints.NONE;
        pEnd.add(new JLabel("CEP:"), gbc);
        gbc.gridx=1; gbc.gridy=0;
        JPanel pCep = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        try {
            MaskFormatter maskCep = new MaskFormatter("#####-###");
            maskCep.setPlaceholderCharacter('_');
            txtCep = new JFormattedTextField(maskCep);
            txtCep.setColumns(6);
        } catch (Exception e) { txtCep = new JFormattedTextField(); }
        txtCep.addFocusListener(new FocusAdapter() { @Override public void focusLost(FocusEvent e) { buscarCep(); }});
        JButton btnCep = new JButton("BUSCAR");
        btnCep.addActionListener(e -> buscarCep());
        pCep.add(txtCep); pCep.add(Box.createHorizontalStrut(5)); pCep.add(btnCep);
        pEnd.add(pCep, gbc);
        gbc.gridx=2; gbc.gridy=0; pEnd.add(new JLabel(" ENDEREÇO: "), gbc);
        gbc.gridx=3; gbc.gridy=0; gbc.weightx=1.0; gbc.fill=GridBagConstraints.HORIZONTAL;
        txtLogradouro = new JTextField(); pEnd.add(txtLogradouro, gbc);
        gbc.gridx=0; gbc.gridy=1; gbc.weightx=0; gbc.fill=GridBagConstraints.NONE;
        pEnd.add(new JLabel("BAIRRO:"), gbc);
        gbc.gridx=1; gbc.gridy=1; gbc.fill=GridBagConstraints.HORIZONTAL;
        txtBairro = new JTextField(); pEnd.add(txtBairro, gbc);
        gbc.gridx=2; gbc.gridy=1; gbc.fill=GridBagConstraints.NONE;
        pEnd.add(new JLabel(" CIDADE/UF: "), gbc);
        gbc.gridx=3; gbc.gridy=1; gbc.fill=GridBagConstraints.HORIZONTAL;
        JPanel pCid = new JPanel(new BorderLayout());
        txtCidade = new JTextField("RIO DE JANEIRO");
        txtUf = new JTextField("RJ"); txtUf.setPreferredSize(new Dimension(30, 20)); txtUf.setHorizontalAlignment(JTextField.CENTER);
        pCid.add(txtCidade, BorderLayout.CENTER); pCid.add(txtUf, BorderLayout.EAST);
        pEnd.add(pCid, gbc);
        gbc.gridx=0; gbc.gridy=2; gbc.gridwidth=4; gbc.fill=GridBagConstraints.HORIZONTAL;
        panelTelefones = new JPanel();
        panelTelefones.setLayout(new BoxLayout(panelTelefones, BoxLayout.Y_AXIS));
        addTelefoneRow();
        JPanel pTelBtn = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton btnAddTel = new JButton("ADICIONAR OUTRO TELEFONE");
        btnAddTel.addActionListener(e -> addTelefoneRow());
        pTelBtn.add(new JLabel("LISTA DE TELEFONES: ")); pTelBtn.add(panelTelefones);
        JPanel pWrapperTel = new JPanel(new BorderLayout());
        pWrapperTel.add(panelTelefones, BorderLayout.CENTER);
        pWrapperTel.add(btnAddTel, BorderLayout.SOUTH);
        pEnd.add(pWrapperTel, gbc);
        gbc.gridwidth=1;
        mainPanel.add(pEnd);

        JPanel pAdmin = criarPainelRustico("DADOS ADMINISTRATIVOS");
        pAdmin.setLayout(new FlowLayout(FlowLayout.LEFT, 20, 5));
        pAdmin.add(new JLabel("MATRÍCULA:")); txtMatricula = new JTextField(15); pAdmin.add(txtMatricula);
        pAdmin.add(new JLabel("Chave de Autorização:")); txtChave = new JTextField(15); pAdmin.add(txtChave);
        mainPanel.add(pAdmin);

        JPanel pOrigem = criarPainelRustico("ORIGEM DO PACIENTE");
        pOrigem.setLayout(new BorderLayout());
        JPanel pSwitch = new JPanel(new FlowLayout(FlowLayout.CENTER));
        ButtonGroup bgOrig = new ButtonGroup();
        rbResidencia = new JRadioButton("VINDO DE RESIDÊNCIA (ELETIVO)", true);
        rbResidencia.setFont(new Font("Tahoma", Font.BOLD, 14));
        rbAmbulancia = new JRadioButton("VINDO DE AMBULÂNCIA (INTERNADO)");
        rbAmbulancia.setFont(new Font("Tahoma", Font.BOLD, 14));
        bgOrig.add(rbResidencia); bgOrig.add(rbAmbulancia);
        pSwitch.add(rbResidencia); pSwitch.add(rbAmbulancia);
        pOrigem.add(pSwitch, BorderLayout.NORTH);
        cardLayout = new CardLayout();
        panelDinamico = new JPanel(cardLayout);
        JPanel cardEletivo = new JPanel();
        cardEletivo.setLayout(new BoxLayout(cardEletivo, BoxLayout.Y_AXIS));
        JPanel pBio = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pBio.add(new JLabel("PESO (KG):")); txtPeso = new JTextField(6); pBio.add(txtPeso);
        pBio.add(new JLabel("ALTURA (M):")); txtAltura = new JTextField(6); pBio.add(txtAltura);
        cardEletivo.add(pBio);
        chkIodo = new JCheckBox("ALERGIA A IODO?");
        JPanel pIodo = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pIodo.add(chkIodo);
        JPanel pIodoDetalhes = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pIodoDetalhes.add(new JLabel("FEZ PREPARO?"));
        rbPreparoSim = new JRadioButton("SIM");
        rbPreparoNao = new JRadioButton("NÃO", true);
        ButtonGroup bgPreparo = new ButtonGroup();
        bgPreparo.add(rbPreparoSim); bgPreparo.add(rbPreparoNao);
        pIodoDetalhes.add(rbPreparoSim); pIodoDetalhes.add(rbPreparoNao);
        pIodo.add(pIodoDetalhes);
        pIodoDetalhes.setVisible(false);
        chkIodo.addActionListener(e -> pIodoDetalhes.setVisible(chkIodo.isSelected()));
        cardEletivo.add(pIodo);
        chkAnti = new JCheckBox("ANTICOAGULANTE?");
        JPanel pAnti = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pAnti.add(chkAnti);
        JPanel pAntiDetalhes = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pAntiDetalhes.add(new JLabel("QUAL?"));
        txtAntiNome = new JTextField(15);
        pAntiDetalhes.add(txtAntiNome);
        chkSusp48 = new JCheckBox("SUSP. 48H?");
        pAntiDetalhes.add(chkSusp48);
        pAnti.add(pAntiDetalhes);
        pAntiDetalhes.setVisible(false);
        chkAnti.addActionListener(e -> pAntiDetalhes.setVisible(chkAnti.isSelected()));
        cardEletivo.add(pAnti);
        chkDiab = new JCheckBox("DIABETES?");
        JPanel pDiab = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pDiab.add(chkDiab);
        JPanel pDiabDetalhes = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pDiabDetalhes.add(new JLabel("MEDICAMENTOS?"));
        txtDiabNome = new JTextField(15);
        pDiabDetalhes.add(txtDiabNome);
        chkSusp24 = new JCheckBox("SUSP. 24H?");
        pDiabDetalhes.add(chkSusp24);
        pDiab.add(pDiabDetalhes);
        pDiabDetalhes.setVisible(false);
        chkDiab.addActionListener(e -> pDiabDetalhes.setVisible(chkDiab.isSelected()));
        cardEletivo.add(pDiab);
        chkCiru = new JCheckBox("CIRURGIA CARDÍACA PRÉVIA?");
        JPanel pCiru = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pCiru.add(chkCiru);
        JPanel pCiruDetalhes = new JPanel(new FlowLayout(FlowLayout.LEFT));
        pCiruDetalhes.add(new JLabel("QUAL?"));
        txtCiruNome = new JTextField(20);
        pCiruDetalhes.add(txtCiruNome);
        pCiru.add(pCiruDetalhes);
        pCiruDetalhes.setVisible(false);
        chkCiru.addActionListener(e -> pCiruDetalhes.setVisible(chkCiru.isSelected()));
        cardEletivo.add(pCiru);
        JPanel cardAmb = new JPanel(new FlowLayout(FlowLayout.LEFT));
        cardAmb.add(new JLabel("HOSPITAL DE ORIGEM:"));
        txtHospitalOrigem = new JTextField(40);
        cardAmb.add(txtHospitalOrigem);
        panelDinamico.add(cardEletivo, "RESIDENCIA");
        panelDinamico.add(cardAmb, "AMBULANCIA");
        pOrigem.add(panelDinamico, BorderLayout.CENTER);
        mainPanel.add(pOrigem);
        ActionListener trocaL = e -> cardLayout.show(panelDinamico, rbResidencia.isSelected() ? "RESIDENCIA" : "AMBULANCIA");
        rbResidencia.addActionListener(trocaL); rbAmbulancia.addActionListener(trocaL);
    }

    private void setupGlassPane() {
        glassPane = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                g.setColor(new Color(0, 0, 0, 100));
                g.fillRect(0, 0, getWidth(), getHeight());
            }
        };
        glassPane.setOpaque(false);
        glassPane.setLayout(new GridBagLayout());
        JLabel message = new JLabel("Volte para a janela para focar");
        message.setFont(new Font("Arial", Font.BOLD, 24));
        message.setForeground(Color.WHITE);
        glassPane.add(message);
        setGlassPane(glassPane);

        addWindowFocusListener(new WindowFocusListener() {
            @Override
            public void windowGainedFocus(WindowEvent e) {
                glassPane.setVisible(false);
            }

            @Override
            public void windowLostFocus(WindowEvent e) {
                glassPane.setVisible(true);
            }
        });
    }


    private void verificarImportacaoArquivo() {
        if (rbAngioplastia.isSelected()) {
            JFileChooser fc = new JFileChooser(outputDirCat);
            fc.setDialogTitle("Selecione a Ficha de Cateterismo Anterior");
            fc.setFileFilter(new FileNameExtensionFilter("Arquivos Word (*.docx)", "docx"));

            if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                importarDadosWord(fc.getSelectedFile());
            }
        }
    }

    private void importarDadosWord(File arquivo) {
        try (FileInputStream fis = new FileInputStream(arquivo);
             XWPFDocument doc = new XWPFDocument(fis)) {

            Map<String, String> dadosExtraidos = new HashMap<>();

            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (int i = 0; i < cells.size(); i++) {
                        String textoCelula = cells.get(i).getText().trim().toUpperCase();
                        verificarEExtrair(textoCelula, "NOME", cells, i, dadosExtraidos);
                        verificarEExtrair(textoCelula, "PACIENTE", cells, i, dadosExtraidos);
                        verificarEExtrair(textoCelula, "NASCIMENTO", cells, i, dadosExtraidos);
                        verificarEExtrair(textoCelula, "NASC", cells, i, dadosExtraidos);
                        verificarEExtrair(textoCelula, "CPF", cells, i, dadosExtraidos);
                        verificarEExtrair(textoCelula, "MÃE", cells, i, dadosExtraidos);
                        verificarEExtrair(textoCelula, "MAE", cells, i, dadosExtraidos);
                        verificarEExtrair(textoCelula, "MATRÍCULA", cells, i, dadosExtraidos);
                    }
                }
            }

            if(dadosExtraidos.containsKey("NOME")) txtNome.setText(dadosExtraidos.get("NOME"));
            else if(dadosExtraidos.containsKey("PACIENTE") && txtNome.getText().isEmpty()) txtNome.setText(dadosExtraidos.get("PACIENTE"));

            if(dadosExtraidos.containsKey("NASCIMENTO")) txtNasc.setText(dadosExtraidos.get("NASCIMENTO"));
            else if(dadosExtraidos.containsKey("NASC")) txtNasc.setText(dadosExtraidos.get("NASC"));

            if(dadosExtraidos.containsKey("CPF")) txtCpf.setText(dadosExtraidos.get("CPF"));

            if(dadosExtraidos.containsKey("MAE")) txtMae.setText(dadosExtraidos.get("MAE"));
            else if(dadosExtraidos.containsKey("MÃE")) txtMae.setText(dadosExtraidos.get("MÃE"));

            if(dadosExtraidos.containsKey("MATRÍCULA")) txtMatricula.setText(dadosExtraidos.get("MATRÍCULA"));

            JOptionPane.showMessageDialog(this, "Dados importados com sucesso!\nVerifique os campos.");

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Erro na importação: " + e.getMessage());
        }
    }

    private void verificarEExtrair(String textoAtual, String chave, List<XWPFTableCell> cells, int indexAtual, Map<String, String> mapa) {
        if (textoAtual.contains(chave)) {
            if (indexAtual + 1 < cells.size()) {
                String valor = cells.get(indexAtual + 1).getText().trim();
                if (!valor.isEmpty()) {
                    mapa.put(chave, valor);
                    return;
                }
            }
            if (textoAtual.contains(":")) {
                String[] partes = textoAtual.split(":");
                if (partes.length > 1) {
                    mapa.put(chave, partes[1].trim());
                }
            }
        }
    }

    private void processarFicha() {
        if (txtNome.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "NOME OBRIGATÓRIO!");
            return;
        }

        File arquivo = gerarWord();
        if (arquivo == null) return;

        try { Desktop.getDesktop().open(arquivo); } catch (Exception ignored) {}

        int enviar = JOptionPane.showConfirmDialog(this, "ENVIAR ESTA FICHA PARA A NUVEM (HEMODINÂMICA)?", "UPLOAD", JOptionPane.YES_NO_OPTION);
        if (enviar == JOptionPane.YES_OPTION) enviarUpload(arquivo);

        if (rascunhoAberto != null) {
            int excluirRascunho = JOptionPane.showConfirmDialog(this, "Esta ficha foi finalizada a partir de um rascunho.\nDeseja excluir o rascunho do servidor?", "Excluir Rascunho?", JOptionPane.YES_NO_OPTION);
            if (excluirRascunho == JOptionPane.YES_OPTION) {
                deletarRascunho(rascunhoAberto);
            }
        }

        int limpar = JOptionPane.showConfirmDialog(this, "LIMPAR TELA PARA O PRÓXIMO?", "LIMPAR", JOptionPane.YES_NO_OPTION);
        if (limpar == JOptionPane.YES_OPTION) {
            limparCampos(this.getContentPane());
            txtNome.requestFocus();
        }
    }

    private File gerarWord() {
        try {
            Map<String, String> dados = coletarDadosDaTela();

            String modeloNome = rbAmbulancia.isSelected() ? "template_internado.docx" : "template_eletivo.docx";
            File modelo = new File(modeloNome);
            if (!modelo.exists()) modelo = new File("templates/" + modeloNome);
            if (!modelo.exists()) {
                File baseDir = new File(outputDirCat).getParentFile();
                if(baseDir != null) modelo = new File(baseDir, "templates/" + modeloNome);
            }

            if (!modelo.exists()) {
                JOptionPane.showMessageDialog(this, "MODELO NÃO ENCONTRADO: " + modeloNome);
                return null;
            }

            FileInputStream fis = new FileInputStream(modelo);
            XWPFDocument doc = new XWPFDocument(fis);

            substituirNoDoc(doc, dados);

            fis.close();

            String pathDestino = rbAngioplastia.isSelected() ? outputDirAngio : outputDirCat;
            File dir = new File(pathDestino);
            if(!dir.exists()) dir.mkdirs();

            String nomeLimpo = txtNome.getText().replaceAll("[^a-zA-Z0-9 \\p{L}\\-]", "").trim();
            if(nomeLimpo.isEmpty()) nomeLimpo = "SEM_NOME";
            String fname = nomeLimpo + ".docx";

            File saida = new File(dir, fname);
            FileOutputStream fos = new FileOutputStream(saida);
            doc.write(fos); fos.close(); doc.close();

            return saida;
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "ERRO: " + e.getMessage());
            return null;
        }
    }

    private void substituirNoDoc(XWPFDocument doc, Map<String, String> dados) {
        for (XWPFParagraph p : doc.getParagraphs()) replaceInParagraph(p, dados);
        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) replaceInParagraph(p, dados);
                }
            }
        }
    }

    private void replaceInParagraph(XWPFParagraph p, Map<String, String> dados) {
        String text = p.getText();
        if (text == null || text.isEmpty()) return;

        boolean found = false;
        for (String key : dados.keySet()) {
            if (text.contains(key)) { found = true; break; }
        }
        if (!found) return;

        String novoTexto = text;
        for (Map.Entry<String, String> entry : dados.entrySet()) {
            if (entry.getValue() != null) {
                novoTexto = novoTexto.replace(entry.getKey(), entry.getValue());
            }
        }

        String fontFamily = null;
        int fontSize = -1;
        boolean isBold = false;
        boolean isItalic = false;

        if (!p.getRuns().isEmpty()) {
            XWPFRun r = p.getRuns().get(0);
            fontFamily = r.getFontFamily();
            fontSize = r.getFontSize();
            isBold = r.isBold();
            isItalic = r.isItalic();
        }

        while (p.getRuns().size() > 0) p.removeRun(0);

        XWPFRun run = p.createRun();
        run.setText(novoTexto);

        if (fontFamily != null) run.setFontFamily(fontFamily);
        if (fontSize != -1) run.setFontSize(fontSize);
        run.setBold(isBold);
        run.setItalic(isItalic);
    }

    private JPanel criarPainelRustico(String titulo) {
        JPanel p = new JPanel();
        p.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEtchedBorder(EtchedBorder.LOWERED),
                titulo, TitledBorder.LEFT, TitledBorder.TOP,
                new Font("Tahoma", Font.BOLD, 12), Color.BLACK
        ));
        return p;
    }

    private void addTelefoneRow() {
        JPanel p = new JPanel(new FlowLayout(FlowLayout.LEFT, 2, 2));
        JFormattedTextField ddd, num;
        try {
            MaskFormatter maskDdd = new MaskFormatter("(##)");
            maskDdd.setPlaceholderCharacter('_');
            ddd = new JFormattedTextField(maskDdd);
            ddd.setColumns(2);
        } catch (Exception e) { ddd = new JFormattedTextField(); }
        try {
            MaskFormatter maskNum = new MaskFormatter("#####-####");
            maskNum.setPlaceholderCharacter('_');
            num = new JFormattedTextField(maskNum);
            num.setColumns(7);
        } catch (Exception e) { num = new JFormattedTextField(); }

        JTextField nom = new JTextField(12);
        p.add(new JLabel("DDD:")); p.add(ddd);
        p.add(new JLabel(" NUM:")); p.add(num);
        p.add(new JLabel(" CONTATO:")); p.add(nom);
        listaPaineisTelefone.add(p);
        panelTelefones.add(p);
        panelTelefones.revalidate();
        SwingUtilities.invokeLater(ddd::requestFocusInWindow);
    }

    private String gerarTxt(String tipo) {
        if(tipo.equals("IODO")) return !chkIodo.isSelected() ? "NÃO" : "SIM (FEZ PREPARO: " + (rbPreparoSim.isSelected()?"SIM":"NÃO") + ")";
        if(tipo.equals("ANTI")) return !chkAnti.isSelected() ? "NÃO" : "SIM - " + txtAntiNome.getText().toUpperCase() + (chkSusp48.isSelected()?" (SUSPENSO 48H)":"");
        if(tipo.equals("DIAB")) return !chkDiab.isSelected() ? "NÃO" : "SIM - " + txtDiabNome.getText().toUpperCase() + (chkSusp24.isSelected()?" (SUSPENSO 24H)":"");
        if(tipo.equals("CIRU")) return !chkCiru.isSelected() ? "NÃO" : "SIM - " + txtCiruNome.getText().toUpperCase() + " (REVASC)";
        return "";
    }

    private void buscarCep() {
        String c = txtCep.getText().replaceAll("[^\\d]", "");
        if(c.length()!=8) return;
        new Thread(()->{
            try {
                URL u = new URL("https://viacep.com.br/ws/"+c+"/json/");
                Scanner s = new Scanner(u.openStream(), "UTF-8").useDelimiter("\\A");
                JSONObject j = new JSONObject(s.next());
                if(!j.has("erro")) SwingUtilities.invokeLater(()->{
                    txtLogradouro.setText(j.optString("logradouro","").toUpperCase());
                    txtBairro.setText(j.optString("bairro","").toUpperCase());
                    txtCidade.setText(j.optString("localidade","").toUpperCase());
                    txtUf.setText(j.optString("uf","").toUpperCase());
                });
            } catch(Exception e){}
        }).start();
    }

    private void enviarUpload(File f) {
        new Thread(() -> {
            try {
                String boundary = "Boundary" + System.currentTimeMillis();
                HttpURLConnection c = (HttpURLConnection) new URL(serverUrl).openConnection();

                c.setDoOutput(true);
                c.setRequestMethod("POST");
                c.setRequestProperty("Content-Type", "multipart/form-data; boundary=" + boundary);
                c.setRequestProperty("User-Agent", "AppRecepcao Java Client");
                c.setConnectTimeout(15000);
                c.setReadTimeout(15000);

                OutputStream o = c.getOutputStream();

                String nomeCodificado = URLEncoder.encode(f.getName(), "UTF-8");

                String header =
                        "--" + boundary + "\r\n" +
                                "Content-Disposition: form-data; name=\"files\"; filename=\"" + nomeCodificado + "\"\r\n" +
                                "Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n";

                o.write(header.getBytes(StandardCharsets.UTF_8));
                Files.copy(f.toPath(), o);
                o.write(("\r\n--" + boundary + "--\r\n").getBytes(StandardCharsets.UTF_8));
                o.flush();
                o.close();

                int cod = c.getResponseCode();

                SwingUtilities.invokeLater(() ->
                        JOptionPane.showMessageDialog(this,
                                cod == 200 ? "SUCESSO NO ENVIO!" : "ERRO NO ENVIO: " + cod)
                );

            } catch (Exception e) {
                SwingUtilities.invokeLater(() ->
                        JOptionPane.showMessageDialog(this, "FALHA CONEXÃO: " + e.getMessage())
                );
            }
        }).start();
    }

    private void abrirDialogoInternado() {
        JTextField nomeField = new JTextField();
        JTextField cnsField = new JTextField();
        JTextField procedenciaField = new JTextField();
        JFormattedTextField nascField = null;
        try {
            nascField = new JFormattedTextField(new MaskFormatter("##/##/####"));
        } catch (Exception e) { /* Ignored */ }

        JPanel panel = new JPanel(new GridLayout(0, 1, 5, 5));
        panel.add(new JLabel("Nome Completo:"));
        panel.add(nomeField);
        panel.add(new JLabel("CNS:"));
        panel.add(cnsField);
        panel.add(new JLabel("Data de Nascimento:"));
        panel.add(nascField);
        panel.add(new JLabel("Procedência (Ex: CTI, Andar):"));
        panel.add(procedenciaField);

        int result = JOptionPane.showConfirmDialog(this, panel, "Lançar Paciente Internado",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        if (result == JOptionPane.OK_OPTION) {
            if (nomeField.getText().trim().isEmpty()) {
                JOptionPane.showMessageDialog(this, "O nome do paciente é obrigatório.", "Erro", JOptionPane.ERROR_MESSAGE);
                return;
            }
            enviarDadosInternado(
                    nomeField.getText(),
                    cnsField.getText(),
                    nascField.getText(),
                    procedenciaField.getText()
            );
        }
    }

    private void enviarDadosInternado(String nome, String cns, String nasc, String procedencia) {
        new Thread(() -> {
            try {
                URL url = new URL(serverUrl + "/api/laudos/internado");
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("POST");
                conn.setRequestProperty("Content-Type", "application/json; utf-8");
                conn.setRequestProperty("Accept", "application/json");
                conn.setDoOutput(true);

                JSONObject json = new JSONObject();
                json.put("nome", nome);
                json.put("cns", cns);
                json.put("nasc", nasc);
                json.put("procedencia", procedencia);

                try (OutputStream os = conn.getOutputStream()) {
                    os.write(json.toString().getBytes(StandardCharsets.UTF_8));
                }

                int code = conn.getResponseCode();
                if (code == 201) {
                    SwingUtilities.invokeLater(() ->
                            JOptionPane.showMessageDialog(this, "Laudos (CAT e ANGIO) gerados com sucesso!", "Sucesso", JOptionPane.INFORMATION_MESSAGE));
                } else {
                    try (BufferedReader br = new BufferedReader(new InputStreamReader(conn.getErrorStream(), StandardCharsets.UTF_8))) {
                        throw new IOException("Erro no servidor: " + code + " - " + br.lines().collect(Collectors.joining("\n")));
                    }
                }

            } catch (Exception e) {
                e.printStackTrace();
                SwingUtilities.invokeLater(() ->
                        JOptionPane.showMessageDialog(this, "Falha ao enviar dados: " + e.getMessage(), "Erro de Conexão", JOptionPane.ERROR_MESSAGE));
            }
        }).start();
    }


    private void limparCampos(Container c) {
        for(Component cp : c.getComponents()) {
            if(cp instanceof JTextField) ((JTextField)cp).setText("");
            else if(cp instanceof JComboBox) ((JComboBox)cp).setSelectedIndex(-1);
            else if(cp instanceof JCheckBox) ((JCheckBox)cp).setSelected(false);
            else if(cp instanceof Container) limparCampos((Container)cp);
        }
        if(c == getContentPane()) {
            rbCateterismo.setSelected(true); rbResidencia.setSelected(true);
            txtCidade.setText("RIO DE JANEIRO"); txtUf.setText("RJ");
            cardLayout.show(panelDinamico, "RESIDENCIA");
            while(listaPaineisTelefone.size()>1) panelTelefones.remove(listaPaineisTelefone.remove(listaPaineisTelefone.size()-1));
            panelTelefones.revalidate(); panelTelefones.repaint();
        }
    }

    private void loadConfig() {
        try {
            File f=new File(CONFIG_FILE);
            if(f.exists()) {
                JSONObject j=new JSONObject(new String(Files.readAllBytes(f.toPath())));
                serverUrl=j.optString("url",serverUrl);
                outputDirCat=j.optString("pathCat", outputDirCat);
                outputDirAngio=j.optString("pathAngio", outputDirAngio);

                if(j.has("path") && !j.has("pathCat")) outputDirCat = j.getString("path");
            }
        } catch(Exception e){}
    }

    private void saveConfig() {
        try {
            JSONObject j=new JSONObject();
            j.put("url",serverUrl);
            j.put("pathCat", outputDirCat);
            j.put("pathAngio", outputDirAngio);
            Files.write(new File(CONFIG_FILE).toPath(), j.toString(4).getBytes());
        } catch(Exception e){}
    }

    private void abrirConfiguracoes() {
        JTextField u=new JTextField(serverUrl);
        JTextField pCat=new JTextField(outputDirCat);
        JTextField pAng=new JTextField(outputDirAngio);

        Object[] message = {
                "URL do Servidor:", u,
                "Pasta Cateterismo:", pCat,
                "Pasta Angioplastia:", pAng
        };

        if(JOptionPane.showConfirmDialog(this, message, "CONFIG", JOptionPane.OK_CANCEL_OPTION)==0) {
            serverUrl=u.getText();
            outputDirCat=pCat.getText();
            outputDirAngio=pAng.getText();
            saveConfig();
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
            } catch (Exception ignored) {}
            new AppRecepcao().setVisible(true);
        });
    }

    // --- NOVOS MÉTODOS PARA FICHAS PRÉ-PRONTAS ---

    private void salvarRascunho() {
        JTextField nomeRascunhoField = new JTextField(txtNome.getText()); // Sugere o nome do paciente
        JFormattedTextField dataField = null;
        try {
            dataField = new JFormattedTextField(new MaskFormatter("##/##/####"));
            dataField.setText(new SimpleDateFormat("dd/MM/yyyy").format(new Date()));
        } catch (Exception e) { /* Ignored */ }

        JPanel panel = new JPanel(new GridLayout(0, 1, 5, 5));
        panel.add(new JLabel("Nome para identificar o rascunho:"));
        panel.add(nomeRascunhoField);
        panel.add(new JLabel("Data do Agendamento:"));
        panel.add(dataField);

        int result = JOptionPane.showConfirmDialog(this, panel, "Salvar Rascunho",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        if (result != JOptionPane.OK_OPTION) return;

        String nomeRascunho = nomeRascunhoField.getText();
        String dataAgendamento = dataField.getText();

        if (nomeRascunho == null || nomeRascunho.trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "O nome do rascunho é obrigatório.", "Erro", JOptionPane.ERROR_MESSAGE);
            return;
        }

        JSONObject dados = coletarDadosDaTelaComoJson();
        dados.put("nome_rascunho", nomeRascunho);
        try {
            Date data = new SimpleDateFormat("dd/MM/yyyy").parse(dataAgendamento);
            dados.put("data_agendamento", new SimpleDateFormat("yyyy-MM-dd").format(data));
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Data inválida. Use o formato DD/MM/AAAA.", "Erro", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        new Thread(() -> {
            try {
                // Se for um rascunho existente, deleta o antigo antes de salvar
                if (rascunhoAberto != null) {
                    deletarRascunho(rascunhoAberto);
                }

                URL url = new URL(serverUrl + "/api/fichas/pre");
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("POST");
                conn.setRequestProperty("Content-Type", "application/json; utf-8");
                conn.setDoOutput(true);

                try (OutputStream os = conn.getOutputStream()) {
                    os.write(dados.toString().getBytes(StandardCharsets.UTF_8));
                }

                if (conn.getResponseCode() == 201) {
                    JOptionPane.showMessageDialog(this, "Rascunho salvo com sucesso!", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
                    this.dispose();
                } else {
                    throw new IOException("Erro no servidor: " + conn.getResponseCode());
                }
            } catch (Exception e) {
                JOptionPane.showMessageDialog(this, "Falha ao salvar rascunho: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
            }
        }).start();
    }

    private void abrirDialogoRascunhos() {
        JDialog dialog = new JDialog(this, "Abrir Ficha Pré-Pronta", true);
        dialog.setLayout(new BorderLayout(10, 10));

        // Painel de filtro
        JPanel filterPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JFormattedTextField dataField = null;
        try {
            dataField = new JFormattedTextField(new MaskFormatter("##/##/####"));
        } catch (Exception e) {}
        
        JButton filterButton = new JButton("Filtrar");
        DefaultListModel<String> listModel = new DefaultListModel<>();
        JList<String> list = new JList<>(listModel);
        
        JFormattedTextField finalDataField = dataField;
        filterButton.addActionListener(e -> {
            String data = finalDataField.getText().replaceAll("[^\\d]", "");
            if (data.length() == 8) {
                try {
                    Date d = new SimpleDateFormat("ddMMyyyy").parse(data);
                    String dataFormatada = new SimpleDateFormat("yyyy-MM-dd").format(d);
                    buscarRascunhos(dataFormatada, listModel);
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(dialog, "Data inválida.", "Erro", JOptionPane.ERROR_MESSAGE);
                }
            } else {
                buscarRascunhos(null, listModel); // Busca todos
            }
        });

        filterPanel.add(new JLabel("Filtrar por Data:"));
        filterPanel.add(dataField);
        filterPanel.add(filterButton);

        // Painel de botões de ação
        JPanel actionPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton btnAbrir = new JButton("Abrir");
        btnAbrir.addActionListener(e -> {
            if (list.getSelectedValue() != null) {
                carregarRascunho(list.getSelectedValue());
                dialog.dispose();
            }
        });
        JButton btnCancelar = new JButton("Cancelar");
        btnCancelar.addActionListener(e -> dialog.dispose());
        actionPanel.add(btnAbrir);
        actionPanel.add(btnCancelar);

        dialog.add(filterPanel, BorderLayout.NORTH);
        dialog.add(new JScrollPane(list), BorderLayout.CENTER);
        dialog.add(actionPanel, BorderLayout.SOUTH);
        
        dialog.pack();
        dialog.setLocationRelativeTo(this);
        
        buscarRascunhos(null, listModel); // Carga inicial
        dialog.setVisible(true);
    }

    private void buscarRascunhos(String data, DefaultListModel<String> listModel) {
        new Thread(() -> {
            try {
                String urlString = serverUrl + "/api/fichas/pre";
                if (data != null && !data.isEmpty()) {
                    urlString += "?data=" + data;
                }
                URL url = new URL(urlString);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");

                if (conn.getResponseCode() == 200) {
                    try (BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
                        JSONArray rascunhos = new JSONArray(br.lines().collect(Collectors.joining("\n")));
                        SwingUtilities.invokeLater(() -> {
                            listModel.clear();
                            for (int i = 0; i < rascunhos.length(); i++) {
                                listModel.addElement(rascunhos.getString(i));
                            }
                        });
                    }
                } else {
                    throw new IOException("Erro ao buscar rascunhos: " + conn.getResponseCode());
                }
            } catch (Exception e) {
                JOptionPane.showMessageDialog(this, "Falha ao buscar rascunhos: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
            }
        }).start();
    }

    private void carregarRascunho(String filename) {
        new Thread(() -> {
            try {
                String encodedFilename = URLEncoder.encode(filename, StandardCharsets.UTF_8.toString());
                URL url = new URL(serverUrl + "/api/fichas/pre/" + encodedFilename);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");

                if (conn.getResponseCode() == 200) {
                    try (BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
                        JSONObject dados = new JSONObject(br.lines().collect(Collectors.joining("\n")));
                        dados.put("nome_rascunho_original", filename);
                        SwingUtilities.invokeLater(() -> new AppRecepcao(true, dados).setVisible(true));
                    }
                } else {
                    throw new IOException("Erro ao carregar rascunho: " + conn.getResponseCode());
                }
            } catch (Exception e) {
                JOptionPane.showMessageDialog(this, "Falha ao carregar rascunho: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
            }
        }).start();
    }
    
    private void deletarRascunho(String filename) {
        new Thread(() -> {
            try {
                String encodedFilename = URLEncoder.encode(filename, StandardCharsets.UTF_8.toString());
                URL url = new URL(serverUrl + "/api/fichas/pre/" + encodedFilename);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("DELETE");
                if (conn.getResponseCode() == 200) {
                    System.out.println("Rascunho " + filename + " deletado com sucesso.");
                } else {
                    System.err.println("Falha ao deletar rascunho " + filename + ". Código: " + conn.getResponseCode());
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }).start();
    }


    private JSONObject coletarDadosDaTelaComoJson() {
        JSONObject dados = new JSONObject();
        dados.put("rbCateterismo", rbCateterismo.isSelected());
        dados.put("comboMedico", comboMedico.getSelectedIndex());
        dados.put("txtNome", txtNome.getText());
        dados.put("txtNasc", txtNasc.getText());
        dados.put("txtRg", txtRg.getText());
        dados.put("txtCpf", txtCpf.getText());
        dados.put("txtMae", txtMae.getText());
        dados.put("txtPai", txtPai.getText());
        dados.put("txtCep", txtCep.getText());
        dados.put("txtLogradouro", txtLogradouro.getText());
        dados.put("txtBairro", txtBairro.getText());
        dados.put("txtCidade", txtCidade.getText());
        dados.put("txtUf", txtUf.getText());
        dados.put("txtMatricula", txtMatricula.getText());
        dados.put("txtChave", txtChave.getText());
        dados.put("rbResidencia", rbResidencia.isSelected());
        dados.put("txtHospitalOrigem", txtHospitalOrigem.getText());
        dados.put("txtPeso", txtPeso.getText());
        dados.put("txtAltura", txtAltura.getText());
        dados.put("chkIodo", chkIodo.isSelected());
        dados.put("rbPreparoSim", rbPreparoSim.isSelected());
        dados.put("chkAnti", chkAnti.isSelected());
        dados.put("txtAntiNome", txtAntiNome.getText());
        dados.put("chkSusp48", chkSusp48.isSelected());
        dados.put("chkDiab", chkDiab.isSelected());
        dados.put("txtDiabNome", txtDiabNome.getText());
        dados.put("chkSusp24", chkSusp24.isSelected());
        dados.put("chkCiru", chkCiru.isSelected());
        dados.put("txtCiruNome", txtCiruNome.getText());
        return dados;
    }
    
    private Map<String, String> coletarDadosDaTela() {
        Map<String, String> dados = new HashMap<>();
        dados.put("{{PROCEDIMENTO}}", rbAngioplastia.isSelected() ? "ANGIOPLASTIA" : "CATETERISMO");
        dados.put("{{NOME}}", txtNome.getText().toUpperCase());
        dados.put("{{NASCIMENTO}}", txtNasc.getText());
        dados.put("{{CEP}}", txtCep.getText());
        dados.put("{{LOGRADOURO}}", txtLogradouro.getText().toUpperCase());
        dados.put("{{BAIRRO}}", txtBairro.getText().toUpperCase());
        dados.put("{{CIDADE}}", txtCidade.getText().toUpperCase());
        dados.put("{{UF}}", txtUf.getText().toUpperCase());
        dados.put("{{RG}}", txtRg.getText());
        dados.put("{{CPF}}", txtCpf.getText());
        dados.put("{{MAE}}", txtMae.getText().toUpperCase());
        dados.put("{{PAI}}", txtPai.getText().toUpperCase());
        dados.put("{{MATRICULA}}", txtMatricula.getText().toUpperCase());
        dados.put("{{CHAVE}}", txtChave.getText());
        dados.put("{{MEDICO}}", comboMedico.getSelectedItem() != null ? comboMedico.getSelectedItem().toString() : "");
        dados.put("{{DATA_HOJE}}", new SimpleDateFormat("dd/MM/yyyy").format(new Date()));

        StringBuilder sbTel = new StringBuilder();
        for(JPanel p : listaPaineisTelefone) {
            Component[] c = p.getComponents();
            String ddd = ((JFormattedTextField)c[1]).getText();
            String num = ((JFormattedTextField)c[3]).getText();
            String nom = ((JTextField)c[5]).getText();
            if(!ddd.isEmpty() && !num.isEmpty()) sbTel.append(ddd).append(" ").append(num).append(nom.isEmpty()?"":" ("+nom+")").append("  /  ");
        }
        dados.put("{{TELEFONES}}", sbTel.toString());

        if (rbAmbulancia.isSelected()) {
            dados.put("{{ORIGEM}}", txtHospitalOrigem.getText().toUpperCase());
        } else {
            dados.put("{{ORIGEM}}", "RESIDÊNCIA");
            dados.put("{{PESO}}", txtPeso.getText() + " Kg");
            dados.put("{{ALTURA}}", txtAltura.getText() + "m");
            dados.put("{{TXT_IODO}}", gerarTxt("IODO"));
            dados.put("{{TXT_ANTICOAG}}", gerarTxt("ANTI"));
            dados.put("{{TXT_DIABETES}}", gerarTxt("DIAB"));
            dados.put("{{TXT_CIRURGIA}}", gerarTxt("CIRU"));
        }
        return dados;
    }


    private void preencherComRascunho(JSONObject dados) {
        rbCateterismo.setSelected(dados.optBoolean("rbCateterismo", true));
        rbAngioplastia.setSelected(!dados.optBoolean("rbCateterismo", true));
        comboMedico.setSelectedIndex(dados.optInt("comboMedico", -1));
        txtNome.setText(dados.optString("txtNome"));
        txtNasc.setText(dados.optString("txtNasc"));
        txtRg.setText(dados.optString("txtRg"));
        txtCpf.setText(dados.optString("txtCpf"));
        txtMae.setText(dados.optString("txtMae"));
        txtPai.setText(dados.optString("txtPai"));
        txtCep.setText(dados.optString("txtCep"));
        txtLogradouro.setText(dados.optString("txtLogradouro"));
        txtBairro.setText(dados.optString("txtBairro"));
        txtCidade.setText(dados.optString("txtCidade"));
        txtUf.setText(dados.optString("txtUf"));
        txtMatricula.setText(dados.optString("txtMatricula"));
        txtChave.setText(dados.optString("txtChave"));
        rbResidencia.setSelected(dados.optBoolean("rbResidencia", true));
        rbAmbulancia.setSelected(!dados.optBoolean("rbResidencia", true));
        cardLayout.show(panelDinamico, rbResidencia.isSelected() ? "RESIDENCIA" : "AMBULANCIA");
        txtHospitalOrigem.setText(dados.optString("txtHospitalOrigem"));
        txtPeso.setText(dados.optString("txtPeso"));
        txtAltura.setText(dados.optString("txtAltura"));
        chkIodo.setSelected(dados.optBoolean("chkIodo"));
        rbPreparoSim.setSelected(dados.optBoolean("rbPreparoSim"));
        chkAnti.setSelected(dados.optBoolean("chkAnti"));
        txtAntiNome.setText(dados.optString("txtAntiNome"));
        chkSusp48.setSelected(dados.optBoolean("chkSusp48"));
        chkDiab.setSelected(dados.optBoolean("chkDiab"));
        txtDiabNome.setText(dados.optString("txtDiabNome"));
        chkSusp24.setSelected(dados.optBoolean("chkSusp24"));
        chkCiru.setSelected(dados.optBoolean("chkCiru"));
        txtCiruNome.setText(dados.optString("txtCiruNome"));
        
        chkIodo.getActionListeners()[0].actionPerformed(new ActionEvent(this, ActionEvent.ACTION_PERFORMED, null));
        chkAnti.getActionListeners()[0].actionPerformed(new ActionEvent(this, ActionEvent.ACTION_PERFORMED, null));
        chkDiab.getActionListeners()[0].actionPerformed(new ActionEvent(this, ActionEvent.ACTION_PERFORMED, null));
        chkCiru.getActionListeners()[0].actionPerformed(new ActionEvent(this, ActionEvent.ACTION_PERFORMED, null));
        
        this.rascunhoAberto = dados.optString("nome_rascunho_original");
        if (this.rascunhoAberto != null) {
            setTitle(getTitle() + " - [Rascunho: " + rascunhoAberto + "]");
        }
    }
}
