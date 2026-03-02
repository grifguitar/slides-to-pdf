package com.slidesconverter;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.ScreenshotType;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import java.util.regex.*;

public class Main extends JFrame {

    record ConversionState(String presentationId, String previewUrl) {
        static Optional<ConversionState> parse(String url) {
            var m = Pattern.compile("/presentation/d/([a-zA-Z0-9_-]+)").matcher(url);
            if (!m.find())
                return Optional.empty();
            var id = m.group(1);
            // /present открывает слайды на весь экран без серых рамок
            return Optional.of(new ConversionState(
                    id,
                    "https://docs.google.com/presentation/d/%s/present".formatted(id)));
        }
    }

    record Slide(int number, Path path) {
    }

    private final JTextField urlField = new JTextField();
    private final JButton btnConvert = new JButton("Конвертировать в PDF");
    private final JProgressBar progress = new JProgressBar(0, 100);
    private final JLabel lblStatus = new JLabel("Готов к работе");
    private final JTextArea logArea = new JTextArea();

    // Ссылки для корректного завершения при закрытии окна
    private volatile SwingWorker<File, String> currentWorker = null;
    private volatile Browser currentBrowser = null;
    private volatile boolean converting = false;

    public Main() {
        super("GoogleSlides2PDF by kgrigoriy for Anna💕 v1.0");
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        setSize(680, 440);
        setLocationRelativeTo(null);
        buildUi();

        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent e) {
                if (converting) {
                    var ans = JOptionPane.showConfirmDialog(
                            Main.this,
                            "Конвертация ещё выполняется. Прервать и выйти?",
                            "Подтверждение", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
                    if (ans != JOptionPane.YES_OPTION)
                        return;
                    // Отменяем worker и сразу закрываем браузер
                    if (currentWorker != null)
                        currentWorker.cancel(true);
                }
                closeBrowserSafely();
                dispose();
                System.exit(0);
            }
        });

        setVisible(true);
    }

    private void buildUi() {
        var urlLabel = new JLabel("Ссылка на Google Slides:");
        urlField.setFont(new Font("Arial", Font.PLAIN, 13));

        btnConvert.setFont(new Font("Arial", Font.BOLD, 13));
        btnConvert.setBackground(new Color(66, 133, 244));
        btnConvert.setForeground(Color.WHITE);
        btnConvert.setFocusPainted(false);

        var topPanel = new JPanel(new BorderLayout(6, 4));
        topPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 6, 10));
        topPanel.add(urlLabel, BorderLayout.NORTH);
        topPanel.add(urlField, BorderLayout.CENTER);
        topPanel.add(btnConvert, BorderLayout.EAST);

        lblStatus.setFont(new Font("Arial", Font.PLAIN, 11));
        progress.setStringPainted(true);
        var midPanel = new JPanel(new BorderLayout(4, 4));
        midPanel.setBorder(BorderFactory.createEmptyBorder(2, 10, 4, 10));
        midPanel.add(lblStatus, BorderLayout.NORTH);
        midPanel.add(progress, BorderLayout.CENTER);

        logArea.setEditable(false);
        logArea.setFont(new Font("Monospaced", Font.PLAIN, 11));
        var scroll = new JScrollPane(logArea);
        scroll.setBorder(BorderFactory.createTitledBorder("Лог"));
        scroll.setPreferredSize(new java.awt.Dimension(680, 300));

        setLayout(new BorderLayout(6, 6));
        add(topPanel, BorderLayout.NORTH);
        add(midPanel, BorderLayout.CENTER);
        add(scroll, BorderLayout.SOUTH);

        btnConvert.addActionListener(_ -> onConvertClicked());
        urlField.addActionListener(_ -> onConvertClicked());
    }

    private void closeBrowserSafely() {
        var b = currentBrowser;
        if (b != null) {
            try {
                b.close();
            } catch (Exception ignored) {
            }
            currentBrowser = null;
        }
    }

    private void onConvertClicked() {
        var raw = urlField.getText().strip();
        if (raw.isBlank()) {
            JOptionPane.showMessageDialog(this, "Введите ссылку на Google Slides",
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
            return;
        }
        var stateOpt = ConversionState.parse(raw);
        if (stateOpt.isEmpty()) {
            JOptionPane.showMessageDialog(this,
                    "Не удалось извлечь ID из ссылки. Ожидается:\nhttps://docs.google.com/presentation/d/<ID>/...",
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
            return;
        }
        startWorker(stateOpt.get());
    }

    private static boolean isChromiumInstalled() {
        var home = System.getProperty("user.home");
        var candidates = new java.util.ArrayList<File>();

        // Windows
        var localAppData = System.getenv("LOCALAPPDATA");
        if (localAppData != null)
            candidates.add(new File(localAppData, "ms-playwright"));

        // macOS
        candidates.add(new File(home, "Library/Caches/ms-playwright"));

        // Linux
        candidates.add(new File(home, ".cache/ms-playwright"));

        for (var playwrightDir : candidates) {
            if (!playwrightDir.isDirectory())
                continue;
            var dirs = playwrightDir.listFiles(
                    f -> f.isDirectory() && f.getName().startsWith("chromium"));
            if (dirs == null)
                continue;
            for (var dir : dirs) {
                if (new File(dir, "chrome-win/chrome.exe").exists())
                    return true;
                if (new File(dir, "chrome-linux/chrome").exists())
                    return true;
                if (new File(dir, "chrome-mac/Chromium.app/Contents/MacOS/Chromium").exists())
                    return true;
            }
        }
        return false;
    }

    private void startWorker(ConversionState state) {
        btnConvert.setEnabled(false);
        urlField.setEnabled(false);
        progress.setValue(0);
        logArea.setText("");

        var worker = new SwingWorker<File, String>() {

            @Override
            protected File doInBackground() throws Exception {
                return runConversion(state);
            }

            @Override
            protected void process(List<String> chunks) {
                for (var msg : chunks) {
                    if (msg.startsWith("PROGRESS:")) {
                        try {
                            progress.setValue(Integer.parseInt(msg.substring(9)));
                        } catch (NumberFormatException ignored) {
                        }
                    } else if (msg.startsWith("STATUS:"))
                        lblStatus.setText(msg.substring(7));
                    else {
                        logArea.append(msg + "\n");
                        logArea.setCaretPosition(logArea.getDocument().getLength());
                    }
                }
            }

            @Override
            protected void done() {
                converting = false;
                currentWorker = null;
                // currentBrowser уже null — закрыт в finally внутри runConversion
                btnConvert.setEnabled(true);
                urlField.setEnabled(true);
                if (isCancelled()) {
                    lblStatus.setText("Отменено");
                    return;
                }
                try {
                    var pdf = get();
                    progress.setValue(100);
                    lblStatus.setText("Готово!");
                    logArea.append("PDF: " + pdf.getAbsolutePath() + "\n");
                    var ans = JOptionPane.showConfirmDialog(Main.this,
                            "PDF создан:\n%s\n\nОткрыть?".formatted(pdf.getAbsolutePath()),
                            "Успех", JOptionPane.YES_NO_OPTION);
                    if (ans == JOptionPane.YES_OPTION)
                        Desktop.getDesktop().open(pdf);
                } catch (Exception ex) {
                    var cause = ex.getCause() != null ? ex.getCause() : ex;
                    var message = cause.getMessage() != null
                            ? cause.getMessage()
                            : cause.getClass().getSimpleName();
                    progress.setValue(0);
                    lblStatus.setText("Ошибка!");
                    logArea.append("ОШИБКА: " + message + "\n");
                    JOptionPane.showMessageDialog(Main.this,
                            message, "Ошибка", JOptionPane.ERROR_MESSAGE);
                }
            }

            void log(String m) {
                publish(m);
            }

            void stat(String m) {
                publish("STATUS:" + m);
            }

            void prog(int v) {
                publish("PROGRESS:" + v);
            }

            private File runConversion(ConversionState state) throws Exception {
                if (isChromiumInstalled()) {
                    log("Chromium уже установлен");
                } else {
                    log("Первый запуск: загрузка Chromium (~130 MB)...");
                    stat("Загрузка Chromium...");
                    installChromium();
                }
                prog(10);

                log("Запуск Playwright...");
                stat("Запуск браузера...");

                try (var playwright = Playwright.create()) {
                    prog(15);

                    currentBrowser = playwright.chromium().launch(
                            new BrowserType.LaunchOptions().setHeadless(true));
                    var browser = currentBrowser;
                    log("Chromium запущен");
                    prog(20);

                    // 1920x1080 — Full HD, соответствует 16:9 Google Slides
                    var context = browser.newContext(
                            new Browser.NewContextOptions().setViewportSize(1920, 1080));
                    var page = context.newPage();

                    stat("Загрузка презентации...");
                    log("URL: " + state.previewUrl());
                    page.navigate(state.previewUrl());

                    // В /present режиме Google Slides рендерит слайды напрямую в DOM
                    // без iframe — ждём появления контейнера слайда или network idle
                    try {
                        page.waitForSelector(
                                ".punch-viewer-container, .punch-filmstrip, [class*='punch-slide']",
                                new Page.WaitForSelectorOptions().setTimeout(12000));
                        log("Контейнер слайдов найден");
                    } catch (Exception e) {
                        // Фоллбэк: ждём network idle если DOM-элемент не найден
                        log("DOM-селектор не сработал, ждём network idle...");
                        try {
                            page.waitForLoadState(com.microsoft.playwright.options.LoadState.NETWORKIDLE,
                                    new Page.WaitForLoadStateOptions().setTimeout(15000));
                        } catch (Exception e2) {
                            log("Network idle timeout, продолжаем");
                        }
                    }
                    // Небольшой буфер на JS-рендер после загрузки сети
                    page.waitForTimeout(600);
                    prog(25);

                    // Инжектируем CSS — скрываем UI-оверлеи
                    injectCss(page);

                    // Даём фокус через JS — mouse.click() в /present переключает слайд вперёд
                    page.evaluate("document.body.focus()");
                    log("Фокус установлен");

                    List<Slide> slides = new ArrayList<>();
                    try {
                        slides = captureSlides(page);
                    } finally {
                        // context и browser закрываются всегда — даже при исключении
                        try {
                            context.close();
                        } catch (Exception ignored) {
                        }
                        try {
                            browser.close();
                        } catch (Exception ignored) {
                        }
                        currentBrowser = null;
                        log("Браузер закрыт");
                    }

                    if (slides.isEmpty())
                        throw new Exception("Не удалось захватить ни одного слайда. " +
                                "Проверьте что презентация открыта для всех по ссылке.");

                    stat("Создание PDF...");
                    prog(85);
                    try {
                        return buildPdf(slides, state.presentationId());
                    } finally {
                        // Temp-файлы удаляются всегда — даже если buildPdf бросил исключение
                        for (var s : slides) {
                            try {
                                Files.deleteIfExists(s.path());
                            } catch (Exception ignored) {
                            }
                        }
                    }
                }
            }

            private void injectCss(Page page) {
                // Скрываем UI-оверлеи Google Slides (/present показывает панели поверх слайда)
                var css = String.join(" ",
                        "* { cursor: none !important; }",
                        ".punch-filmstrip, .punch-filmstrip-container,",
                        ".punch-viewer-scrubbar-container, .punch-viewer-nav-bars,",
                        ".goog-toolbar, [class*='progress-bar'] { display: none !important; }",
                        "body { background: #000 !important; overflow: hidden !important; }");
                try {
                    page.addStyleTag(new Page.AddStyleTagOptions().setContent(css));
                    log("CSS инжектирован");
                } catch (Exception e) {
                    var msg = e.getMessage() != null ? e.getMessage() : e.getClass().getSimpleName();
                    log("CSS не применён: " + msg);
                }
            }

            private List<Slide> captureSlides(Page page) throws Exception {
                List<Slide> slides = new ArrayList<>();
                byte[] prev = null;
                int sameCount = 0;

                for (int num = 1; num <= 300; num++) {
                    stat("Слайд %d...".formatted(num));

                    // Ждём стабилизации кадра: делаем два скриншота с паузой 150мс
                    // Если они одинаковые — слайд отрендерен, можно снимать
                    byte[] a, b;
                    int attempts = 0;
                    do {
                        page.waitForTimeout(120);
                        a = page.screenshot(new Page.ScreenshotOptions().setType(ScreenshotType.PNG));
                        page.waitForTimeout(80);
                        b = page.screenshot(new Page.ScreenshotOptions().setType(ScreenshotType.PNG));
                        attempts++;
                    } while (!Arrays.equals(a, b) && attempts < 4);
                    // Финальный скриншот после стабилизации
                    byte[] bytes = b;

                    if (prev != null && Arrays.equals(bytes, prev)) {
                        sameCount++;
                        if (sameCount >= 2) {
                            log("Конец презентации, слайдов: " + (num - sameCount));
                            break;
                        }
                    } else {
                        sameCount = 0;
                        var path = Files.createTempFile("slide_%03d_".formatted(num), ".png");
                        Files.write(path, bytes);
                        slides.add(new Slide(num, path));
                    }

                    prev = bytes;
                    prog(25 + Math.min((int) (num / 200.0 * 55), 55));
                    page.keyboard().press("ArrowRight");
                }
                return slides;
            }

            private File buildPdf(List<Slide> slides, String id) throws Exception {
                // Спрашиваем путь сохранения на EDT, ждём ответа
                var outFile = new java.util.concurrent.atomic.AtomicReference<File>();
                var latch = new java.util.concurrent.CountDownLatch(1);

                SwingUtilities.invokeLater(() -> {
                    var desktop = new File(System.getProperty("user.home"), "Desktop");
                    var chooser = new JFileChooser(
                            desktop.exists() ? desktop : new File(System.getProperty("user.home")));
                    chooser.setDialogTitle("Сохранить PDF");
                    chooser.setSelectedFile(new File(chooser.getCurrentDirectory(), id + "_slides.pdf"));
                    chooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("PDF файлы", "pdf"));
                    if (chooser.showSaveDialog(Main.this) == JFileChooser.APPROVE_OPTION) {
                        var f = chooser.getSelectedFile();
                        // Добавляем .pdf если не указано
                        if (!f.getName().toLowerCase().endsWith(".pdf"))
                            f = new File(f.getAbsolutePath() + ".pdf");
                        outFile.set(f);
                    }
                    latch.countDown();
                });

                latch.await();

                if (outFile.get() == null)
                    throw new Exception("Сохранение отменено пользователем");

                try (var doc = new PDDocument()) {
                    for (var slide : slides) {
                        BufferedImage img = ImageIO.read(slide.path().toFile());
                        float w = img.getWidth(), h = img.getHeight();
                        var pg = new PDPage(new PDRectangle(w, h));
                        doc.addPage(pg);
                        var pdImg = PDImageXObject.createFromFile(slide.path().toString(), doc);
                        try (var cs = new PDPageContentStream(doc, pg)) {
                            cs.drawImage(pdImg, 0, 0, w, h);
                        }
                    }
                    doc.save(outFile.get());
                    return outFile.get();
                }
            }

            private void installChromium() throws Exception {
                var jarPath = new File(
                        Main.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getAbsolutePath();
                var javaExe = ProcessHandle.current().info().command().orElse("java");
                var pb = new ProcessBuilder(
                        javaExe, "-cp", jarPath,
                        "com.microsoft.playwright.CLI", "install", "chromium");
                pb.redirectErrorStream(true);
                var proc = pb.start();
                try (var br = new BufferedReader(new InputStreamReader(proc.getInputStream()))) {
                    String line;
                    while ((line = br.readLine()) != null)
                        log(line);
                }
                int code = proc.waitFor();
                if (code != 0)
                    throw new Exception("Установка Chromium завершилась с кодом: " + code);
                log("Chromium установлен");
            }
        };

        currentWorker = worker;
        converting = true;
        worker.execute();
    }

    public static void main(String[] args) throws Exception {
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        SwingUtilities.invokeLater(Main::new);
    }
}
