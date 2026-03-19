import javax.swing.*;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.*;
import java.nio.charset.StandardCharsets;

public class LauncherRecepcao {

    // --- CONFIGURAÇÃO ---
    private static final String SERVER_URL = "https://hemodinamica.souzadev.software";
    private static final String APP_JAR_NAME = "AppRecepcao.jar"; // Nome do jar principal
    private static final String VERSION_FILE = "version_recepcao.dat"; // Arquivo local de versão
    private static final String API_VERSION = "/api/recepcao/version";
    private static final String API_DOWNLOAD = "/api/download/recepcao";

    public static void main(String[] args) {
        System.setProperty("https.protocols", "TLSv1.2"); // Fix Java 8

        // Inicia o app diretamente sem verificar atualizações
        launchApp();
    }

    private static void checkForUpdates(JLabel statusLabel) {
        try {
            String localVersion = "0.0.0";
            File vFile = new File(VERSION_FILE);
            if (vFile.exists()) {
                localVersion = new String(Files.readAllBytes(vFile.toPath()), StandardCharsets.UTF_8).trim();
            }

            statusLabel.setText("Contatando servidor...");
            String jsonResponse = httpGet(SERVER_URL + API_VERSION);
            String serverVersion = extractJsonValue(jsonResponse, "version");

            if (!localVersion.equals(serverVersion)) {
                statusLabel.setText("Atualizando para versão " + serverVersion + "...");

                File tempFile = new File(APP_JAR_NAME + ".tmp");
                downloadFile(SERVER_URL + API_DOWNLOAD, tempFile);

                File currentJar = new File(APP_JAR_NAME);
                if (currentJar.exists()) currentJar.delete();
                tempFile.renameTo(currentJar);

                Files.write(vFile.toPath(), serverVersion.getBytes(StandardCharsets.UTF_8));

                statusLabel.setText("Atualizado!");
                Thread.sleep(1000);
            }
        } catch (Exception e) {
            System.err.println("Erro update: " + e.getMessage());
        }
    }

    private static void launchApp() {
        try {
            File appJar = new File(APP_JAR_NAME);
            if (!appJar.exists()) {
                JOptionPane.showMessageDialog(null, "Erro: " + APP_JAR_NAME + " não encontrado.");
                return;
            }
            ProcessBuilder pb = new ProcessBuilder("java", "-jar", APP_JAR_NAME);
            pb.inheritIO();
            pb.start();
            System.exit(0);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Erro ao iniciar: " + e.getMessage());
        }
    }

    private static String httpGet(String urlStr) throws IOException {
        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0 Launcher");
        conn.setConnectTimeout(5000);
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) result.append(line);
            return result.toString();
        }
    }

    private static void downloadFile(String urlStr, File dest) throws IOException {
        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestProperty("User-Agent", "Mozilla/5.0 Launcher");
        try (InputStream in = conn.getInputStream()) {
            Files.copy(in, dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }
    }

    private static String extractJsonValue(String json, String key) {
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("\"" + key + "\":\\s*\"(.*?)\"").matcher(json);
        return m.find() ? m.group(1) : "";
    }
}