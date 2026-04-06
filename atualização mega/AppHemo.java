import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import javax.swing.border.EmptyBorder;
import javax.sound.sampled.*;

import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

/**
 * AppHemo - Versão Mega Atualização (Quarentena)
 * Foco: Suporte Total a Controles de Conteúdo (SDT) e Tags.
 */
public class AppHemo extends JFrame {

    private static final String CURRENT_VERSION = "11.0-MEGA";
    private String serverUrl = "http://localhost:2400";
    private String iaServerUrl = "http://localhost:8000";    
    private String rootPath = "A:\\projetos atuais\\TRABALHOVOTICOR\\laudos";
    private static final String CONFIG_FILE = "config.json";

    private DefaultTableModel modelEmSala, modelAguardando, modelProntos;
    private JTable tableEmSala, tableAguardando, tableProntos;
    private JLabel statusLabel;
    private JButton btnRefresh;

    private TargetDataLine audioLine;
    private File audioFile;
    private boolean isRecording = false;
    private long startTime;

    public static void main(String[] args) {
        System.setProperty("https.protocols", "TLSv1.2");
        SwingUtilities.invokeLater(() -> new AppHemo().setVisible(true));
    }

    public AppHemo() {
        super("VOTICOR - MEGA v" + CURRENT_VERSION);
        loadConfig(); setupUI(); refreshAll();
    }

    private void setupUI() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); setSize(1400, 850); setLocationRelativeTo(null);
        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); } catch (Exception ignored) {}
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10)); mainPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
        JPanel headerPanel = new JPanel(new BorderLayout());
        JLabel titleLabel = new JLabel("Painel de Controle Médico - Hemodinâmica"); titleLabel.setFont(new Font("SansSerif", Font.BOLD, 22));
        JPanel topBtns = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton btnConfig = new JButton("⚙ Config"); btnConfig.addActionListener(e -> showConfigDialog());
        btnRefresh = new JButton("↻ Sincronizar Agora"); btnRefresh.addActionListener(e -> refreshAll());
        topBtns.add(btnConfig); topBtns.add(btnRefresh);
        headerPanel.add(titleLabel, BorderLayout.WEST); headerPanel.add(topBtns, BorderLayout.EAST);
        mainPanel.add(headerPanel, BorderLayout.NORTH);
        JPanel columnsPanel = new JPanel(new GridLayout(1, 3, 15, 0));
        columnsPanel.add(createColumnPanel("EM SALA", modelEmSala = new DefaultTableModel(new String[]{"Paciente", "Início", "Status"}, 0), tableEmSala = new JTable(), new Color(255, 248, 225)));
        columnsPanel.add(createColumnPanel("AGUARDANDO LAUDO", modelAguardando = new DefaultTableModel(new String[]{"Paciente", "Arquivo", "Tipo", "CNS", "Nasc", "Procedência"}, 0), tableAguardando = new JTable(), new Color(232, 245, 233)));
        columnsPanel.add(createColumnPanel("LAUDO PRONTO", modelProntos = new DefaultTableModel(new String[]{"Paciente/Arquivo"}, 0), tableProntos = new JTable(), new Color(236, 239, 241)));
        mainPanel.add(columnsPanel, BorderLayout.CENTER);
        statusLabel = new JLabel("Operacional."); mainPanel.add(statusLabel, BorderLayout.SOUTH);
        add(mainPanel); setupTableEvents();
    }

    private JPanel createColumnPanel(String title, DefaultTableModel model, JTable table, Color bgColor) {
        JPanel p = new JPanel(new BorderLayout(5, 5)); p.setBackground(bgColor); p.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(200, 200, 200)), new EmptyBorder(5, 5, 5, 5)));
        JLabel lbl = new JLabel(title, SwingConstants.CENTER); lbl.setFont(new Font("SansSerif", Font.BOLD, 15)); p.add(lbl, BorderLayout.NORTH);
        table.setModel(model); table.setRowHeight(40); JScrollPane scroll = new JScrollPane(table); scroll.getViewport().setBackground(Color.WHITE); p.add(scroll, BorderLayout.CENTER);
        JButton btn = title.contains("EM SALA") ? new JButton("Abrir Tablet") : title.contains("AGUARDANDO") ? new JButton("ELABORAR LAUDO") : new JButton("Ver Documento");
        if(title.contains("AGUARDANDO")) { btn.setBackground(new Color(40, 167, 69)); btn.setForeground(Color.WHITE); }
        btn.addActionListener(e -> { 
            if(title.contains("EM SALA")) { try { Desktop.getDesktop().browse(new java.net.URI(serverUrl + "/tablet")); } catch (Exception ignored) {} }
            else if(title.contains("AGUARDANDO")) { int r = tableAguardando.getSelectedRow(); if (r != -1) handleAperturaLaudo(r); }
            else { int r = tableProntos.getSelectedRow(); if (r != -1) { String arq = (String) modelProntos.getValueAt(r, 0); try { Desktop.getDesktop().browse(new java.net.URI(serverUrl + "/view/" + encodePath(arq))); } catch (Exception ignored) {} } }
        });
        p.add(btn, BorderLayout.SOUTH); return p;
    }

    private void setupTableEvents() { tableAguardando.addMouseListener(new MouseAdapter() { public void mouseClicked(MouseEvent e) { if (e.getClickCount() == 2) { int r = tableAguardando.getSelectedRow(); if (r != -1) handleAperturaLaudo(r); } } }); }

    private void handleAperturaLaudo(int row) {
        String pac = (String) modelAguardando.getValueAt(row, 0);
        String arq = (String) modelAguardando.getValueAt(row, 1);
        String tipo = (String) modelAguardando.getValueAt(row, 2);
        
        Map<String, String> map = new HashMap<>();
        String nomePC = toPascalCase(pac);
        String procPC = toPascalCase((String) modelAguardando.getValueAt(row, 5));
        String nasc = (String) modelAguardando.getValueAt(row, 4);
        String cns = (String) modelAguardando.getValueAt(row, 3);
        String dataH = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));

        // Mapeamento para Tags {{}} e para IDs de Controles de Conteúdo
        map.put("{{NOME}}", nomePC); map.put("campo_nome", nomePC);
        map.put("{{PROCEDENCIA}}", procPC); map.put("{{PROCEDÊNCIA}}", procPC); map.put("campo_procedencia", procPC);
        map.put("{{NASC}}", nasc); map.put("{{NASCIMENTO}}", nasc); map.put("campo_nasc", nasc);
        map.put("{{CNS}}", cns); map.put("campo_cns", cns);
        map.put("{{DATA_HOJE}}", dataH); map.put("campo_data", dataH);

        Object[] options = {"⌨ Digitar", "🎙 Ditar (IA)", "Cancelar"};
        int choice = JOptionPane.showOptionDialog(this, "Paciente: " + pac, "Modo", JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, options[0]);
        if (choice == 0) abrirLaudoManual(arq, map);
        else if (choice == 1) showGravadorIA(pac, arq, tipo, map);
    }

    private String toPascalCase(String val) {
        if (val == null || val.isEmpty() || val.equalsIgnoreCase("NomeNaoIdentificado")) return val;
        StringBuilder sb = new StringBuilder(); boolean next = true;
        for (char c : val.toLowerCase().toCharArray()) { if (Character.isSpaceChar(c)) { next = true; sb.append(c); } else if (next) { sb.append(Character.toUpperCase(c)); next = false; } else sb.append(c); }
        return sb.toString().trim();
    }

    private void showGravadorIA(String pac, String arqO, String tipo, Map<String, String> map) {
        JDialog dialog = new JDialog(this, "🎙 IA - " + pac, true);
        dialog.setLayout(new BorderLayout(20, 20)); dialog.setSize(400, 300); dialog.setLocationRelativeTo(this);
        JLabel lblTime = new JLabel("00:00", SwingConstants.CENTER); lblTime.setFont(new Font("Monospaced", Font.BOLD, 48));
        JPanel btnPanel = new JPanel(new FlowLayout());
        JButton btnRec = new JButton("🔴 INICIAR"); JButton btnStop = new JButton("⬛ PARAR"); JButton btnSend = new JButton("🚀 ENVIAR");
        btnStop.setEnabled(false); btnSend.setEnabled(false);
        javax.swing.Timer timer = new javax.swing.Timer(1000, e -> { long elapsed = (System.currentTimeMillis() - startTime) / 1000; lblTime.setText(String.format("%02d:%02d", elapsed / 60, elapsed % 60)); });
        btnRec.addActionListener(e -> { if (startRecording()) { startTime = System.currentTimeMillis(); timer.start(); btnRec.setEnabled(false); btnStop.setEnabled(true); } });
        btnStop.addActionListener(e -> { stopRecording(); timer.stop(); btnStop.setEnabled(false); btnSend.setEnabled(true); });
        btnSend.addActionListener(e -> { dialog.dispose(); processarIA(pac, arqO, tipo, map); });
        dialog.add(lblTime, BorderLayout.CENTER); dialog.add(btnPanel, BorderLayout.SOUTH); btnPanel.add(btnRec); btnPanel.add(btnStop); btnPanel.add(btnSend); dialog.setVisible(true);
    }

    private void processarIA(String pac, String arqO, String tipo, Map<String, String> map) {
        new Thread(() -> {
            try {
                setStatus("Processando IA...");
                String template = tipo.equals("ANGIO") ? "Laudo de Angioplastia_ditado.docx" : "Laudo de cateterismo_ditado.docx";
                JSONObject resp = uploadAudioToIA(audioFile, template);
                if (resp.getString("status").equals("success")) {
                    String num = calcularProximoNumero();
                    String numC = JOptionPane.showInputDialog(this, "Número do Exame:", num); if (numC == null) return;
                    map.put("{{NUM_EXAME}}", numC); map.put("campo_num_exame", numC);
                    File pasta = criarEstruturaDePastas(numC, pac, tipo.equals("ANGIO"));
                    File finalFile = new File(pasta, pac.replaceAll("[^a-zA-Z0-9]", "") + (tipo.equals("ANGIO") ? " PTCA.docx" : ".docx"));
                    downloadFile(iaServerUrl + resp.getString("download_url"), finalFile);
                    preencherDocxTotal(finalFile, map);
                    notifyServerDone(arqO); refreshAll(); abrirArquivo(finalFile);
                }
            } catch (Exception e) { JOptionPane.showMessageDialog(this, "Erro IA: " + e.getMessage()); }
        }).start();
    }

    private void abrirLaudoManual(String arq, Map<String, String> map) {
        new Thread(() -> {
            try {
                String num = calcularProximoNumero();
                String numC = JOptionPane.showInputDialog(this, "Número do Exame:", num); if (numC == null) return;
                map.put("{{NUM_EXAME}}", numC); map.put("campo_num_exame", numC);
                File temp = File.createTempFile("manual_", ".docx");
                downloadFile(serverUrl + "/api/baixar/" + encodePath(arq), temp);
                boolean ptca = arq.toUpperCase().contains("PTCA");
                File pasta = criarEstruturaDePastas(numC, (String)map.get("{{NOME}}"), ptca);
                File finalFile = new File(pasta, ((String)map.get("{{NOME}}")).replaceAll("[^a-zA-Z0-9]", "") + (ptca ? " PTCA.docx" : ".docx"));
                try (FileInputStream fis = new FileInputStream(temp); XWPFDocument doc = new XWPFDocument(fis)) {
                    processarElementosDoc(doc, map);
                    try (FileOutputStream fos = new FileOutputStream(finalFile)) { doc.write(fos); }
                }
                temp.delete(); notifyServerDone(arq); refreshAll(); abrirArquivo(finalFile);
            } catch (Exception e) { e.printStackTrace(); }
        }).start();
    }

    private void preencherDocxTotal(File file, Map<String, String> dados) throws Exception {
        try (FileInputStream fis = new FileInputStream(file); XWPFDocument doc = new XWPFDocument(fis)) {
            processarElementosDoc(doc, dados);
            try (FileOutputStream fos = new FileOutputStream(file)) { doc.write(fos); }
        }
    }

    private void processarElementosDoc(XWPFDocument doc, Map<String, String> dados) {
        // IBody permite acessar Parágrafos, Tabelas e SDTs (Controles de Conteúdo)
        for (IBodyElement elem : doc.getBodyElements()) {
            if (elem instanceof XWPFParagraph) { substituirNoParagrafo((XWPFParagraph) elem, dados); }
            else if (elem instanceof XWPFTable) { processarTabela((XWPFTable) elem, dados); }
            else if (elem instanceof XWPFSDT) { processarSDT((XWPFSDT) elem, dados); }
        }
    }

    private void processarSDT(XWPFSDT sdt, Map<String, String> dados) {
        // Tenta preencher pelo Tag/Título do Controle de Conteúdo
        String tag = sdt.getTag();
        String title = sdt.getTitle();
        for (Map.Entry<String, String> entry : dados.entrySet()) {
            if (entry.getKey().equalsIgnoreCase(tag) || entry.getKey().equalsIgnoreCase(title)) {
                // Injeta o texto no controle de conteúdo
                sdt.getSdtContent().setText(entry.getValue());
                return;
            }
        }
        // Se não achou pela Tag, procura {{tags}} dentro do texto do SDT
        String text = sdt.getContent().getText();
        if (text != null) {
            for (Map.Entry<String, String> entry : dados.entrySet()) {
                if (text.contains(entry.getKey())) {
                    sdt.getSdtContent().setText(text.replace(entry.getKey(), entry.getValue()));
                }
            }
        }
    }

    private void processarTabela(XWPFTable t, Map<String, String> dados) {
        for (XWPFTableRow row : t.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) { substituirNoParagrafo(p, dados); }
                // Tabelas podem ter SDTs dentro das células
                for (XWPFSDT sdt : cell.getBodyElements().stream().filter(e -> e instanceof XWPFSDT).map(e -> (XWPFSDT)e).toList()) {
                    processarSDT(sdt, dados);
                }
            }
        }
    }

    private void substituirNoParagrafo(XWPFParagraph p, Map<String, String> dados) {
        String text = p.getText(); boolean mod = false;
        if (text == null) return;
        for (Map.Entry<String, String> e : dados.entrySet()) {
            if (text.contains(e.getKey())) { text = text.replace(e.getKey(), e.getValue() != null ? e.getValue() : ""); mod = true; }
        }
        if (mod) {
            if (p.getRuns().size() > 0) {
                XWPFRun f = p.getRuns().get(0); String font = f.getFontFamily(); int size = f.getFontSize(); boolean b = f.isBold();
                for (int i=p.getRuns().size()-1; i>=0; i--) p.removeRun(i);
                XWPFRun nr = p.createRun(); nr.setText(text);
                if (font != null) nr.setFontFamily(font); if (size != -1) nr.setFontSize(size); nr.setBold(b);
            } else { p.createRun().setText(text); }
        }
    }

    private void refreshAll() { new Thread(() -> { try { updateEmSala(); updateAguardando(); updateProntos(); setStatus("Sincronizado."); } catch (Exception e) { setStatus("Erro conexão."); } }).start(); }
    private void updateEmSala() throws Exception { String j = httpGet(serverUrl + "/api/sala/ativo"); JSONObject o = new JSONObject(j); SwingUtilities.invokeLater(() -> { modelEmSala.setRowCount(0); if (o.has("nome")) { String proc = o.getString("procedimento").toUpperCase(); if (o.optBoolean("evoluiu_angioplastia", false)) proc = "CAT + ANGIO"; modelEmSala.addRow(new Object[]{o.getString("nome"), o.getString("inicio"), proc}); } }); }
    private void updateAguardando() throws Exception { String j = httpGet(serverUrl + "/api/pendentes"); JSONArray arr = new JSONArray(j); SwingUtilities.invokeLater(() -> { modelAguardando.setRowCount(0); for (int i = 0; i < arr.length(); i++) { JSONObject o = arr.getJSONObject(i); modelAguardando.addRow(new Object[]{o.getString("nome"), o.getString("arquivo"), o.getString("arquivo").contains("PTCA") ? "ANGIO" : "CAT", o.optString("cns", ""), o.optString("nasc", ""), o.optString("procedencia", "")}); } }); }
    private void updateProntos() throws Exception { String j = httpGet(serverUrl + "/api/historico"); JSONArray arr = new JSONArray(j); SwingUtilities.invokeLater(() -> { modelProntos.setRowCount(0); for (int i = 0; i < arr.length(); i++) modelProntos.addRow(new Object[]{arr.getJSONObject(i).getString("arquivo")}); }); }
    private boolean startRecording() { try { AudioFormat f = new AudioFormat(44100, 16, 1, true, false); audioLine = (TargetDataLine) AudioSystem.getLine(new DataLine.Info(TargetDataLine.class, f)); audioLine.open(f); audioLine.start(); audioFile = File.createTempFile("rec_", ".wav"); new Thread(() -> { try (AudioInputStream ais = new AudioInputStream(audioLine)) { AudioSystem.write(ais, AudioFileFormat.Type.WAVE, audioFile); } catch (Exception ignored) {} }).start(); return true; } catch (Exception ex) { return false; } }
    private void stopRecording() { if (audioLine != null) { audioLine.stop(); audioLine.close(); } }
    private String calcularProximoNumero() { int m = buscarMaiorNumeroEmMes(LocalDate.now()); if (m == 0) m = buscarMaiorNumeroEmMes(LocalDate.now().minusMonths(1)); if (m == 0) return "42.001"; int p = m + 1; String s = String.format("%05d", p); return s.substring(0, 2) + "." + s.substring(2); }
    private int buscarMaiorNumeroEmMes(LocalDate d) { File pasta = getPastaMes(d); List<File> locais = Arrays.asList(pasta, new File(pasta, "PACIENTE")); int maior = 0; java.util.regex.Pattern pat = java.util.regex.Pattern.compile("(\\d{2})\\.(\\d{3})"); for (File dir : locais) { if (dir.exists() && dir.isDirectory()) { File[] files = dir.listFiles(); if (files != null) { for (File f : files) { if (f.isDirectory()) { java.util.regex.Matcher m = pat.matcher(f.getName()); if (m.find()) { try { int num = Integer.parseInt(m.group(1) + m.group(2)); if (num > maior) maior = num; } catch (Exception ignored) {} } } } } } } return maior; }
    private File getPastaMes(LocalDate d) { String[] m = {"", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"}; return new File(rootPath, String.format("%02d %s", d.getMonthValue(), m[d.getMonthValue()])); }
    private File criarEstruturaDePastas(String num, String nomeP, boolean ptca) { File pMes = getPastaMes(LocalDate.now()); if (!pMes.exists()) pMes.mkdirs(); File pPac = new File(pMes, "PACIENTE"); if (!pPac.exists()) pPac.mkdirs(); File f = new File(pPac, num + " " + nomeP.replaceAll("[^a-zA-Z0-9 ]", "") + (ptca ? " PTCA" : "")); f.mkdirs(); return f; }
    private String extrairNomeDoArquivo(String n) { try { String s = n.substring(0, n.lastIndexOf('.')); String[] p = s.split("_"); if (p.length >= 2) { StringBuilder sb = new StringBuilder(); for (int i = 1; i < p.length; i++) { if (p[i].equalsIgnoreCase("CATETERISMO") || p[i].equalsIgnoreCase("PTCA") || p[i].equalsIgnoreCase("ANGIOPLASTIA")) break; if (sb.length() > 0) sb.append(" "); sb.append(p[i]); } return sb.toString().trim(); } } catch (Exception e) {} return "Paciente"; }
    private JSONObject uploadAudioToIA(File f, String t) throws Exception { String b = "---" + System.currentTimeMillis(); HttpURLConnection c = (HttpURLConnection) new URL(iaServerUrl + "/process").openConnection(); c.setDoOutput(true); c.setRequestMethod("POST"); c.setRequestProperty("Content-Type", "multipart/form-data; boundary=" + b); try (OutputStream out = c.getOutputStream(); PrintWriter w = new PrintWriter(new OutputStreamWriter(out, "UTF-8"), true)) { w.append("--").append(b).append("\r\n").append("Content-Disposition: form-data; name=\"template_name\"\r\n\r\n").append(t).append("\r\n"); w.append("--").append(b).append("\r\n").append("Content-Disposition: form-data; name=\"audio_file\"; filename=\"").append(f.getName()).append("\"\r\n").append("Content-Type: audio/wav\r\n\r\n"); w.flush(); Files.copy(f.toPath(), out); out.flush(); w.append("\r\n--").append(b).append("--\r\n").flush(); } try (Scanner s = new Scanner(c.getInputStream())) { s.useDelimiter("\\A"); return new JSONObject(s.hasNext() ? s.next() : "{}"); } }
    private String httpGet(String u) throws IOException { HttpURLConnection h = (HttpURLConnection) new URL(u).openConnection(); h.setConnectTimeout(5000); if (h.getResponseCode() >= 400) return "{}"; try (Scanner s = new Scanner(h.getInputStream())) { s.useDelimiter("\\A"); return s.hasNext() ? s.next() : "{}"; } }
    private void downloadFile(String u, File d) throws IOException { try (InputStream in = new URL(u).openStream()) { Files.copy(in, d.toPath(), StandardCopyOption.REPLACE_EXISTING); } }
    private void abrirArquivo(File f) { if (f != null && f.exists()) { try { Desktop.getDesktop().open(f); } catch (Exception ignored) {} } }
    private void notifyServerDone(String fn) { new Thread(() -> { try { HttpURLConnection c = (HttpURLConnection) new URL(serverUrl + "/api/concluir/" + encodePath(fn)).openConnection(); c.setRequestMethod("POST"); c.getResponseCode(); } catch (Exception ignored) {} }).start(); }
    private String encodePath(String s) { try { return URLEncoder.encode(s, "UTF-8").replace("+", "%20"); } catch (Exception e) { return s; } }
    private void loadConfig() { try { File f = new File(CONFIG_FILE); if (f.exists()) { JSONObject j = new JSONObject(new String(Files.readAllBytes(f.toPath()))); serverUrl = j.optString("url", serverUrl); rootPath = j.optString("path", rootPath); } } catch (Exception ignored) {} }
    private void showConfigDialog() { JTextField u = new JTextField(serverUrl); JTextField p = new JTextField(rootPath); Object[] m = {"URL:", u, "Pasta:", p}; if (JOptionPane.showConfirmDialog(this, m, "Config", JOptionPane.OK_CANCEL_OPTION) == JOptionPane.OK_OPTION) { serverUrl = u.getText().trim(); rootPath = p.getText().trim(); saveConfig(); refreshAll(); } }
    private void saveConfig() { try { JSONObject j = new JSONObject(); j.put("url", serverUrl); j.put("path", rootPath); Files.write(Paths.get(CONFIG_FILE), j.toString(4).getBytes()); } catch (Exception ignored) {} }
    private void setStatus(String m) { SwingUtilities.invokeLater(() -> statusLabel.setText(m)); }
}
