package com.zlz.split.file.splitfile.service;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.io.RandomAccessReadBuffer;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.regex.Pattern;

public class FileWordCounter {

    // 匹配中文字符
    private static final Pattern CHINESE_PATTERN = Pattern.compile("[\\u4e00-\\u9fa5]");
    // 匹配英文字母
    private static final Pattern LETTER_PATTERN = Pattern.compile("[a-zA-Z]");

    /**
     * 统计文件中的“字数”（中文字符 + 英文字母），忽略图片
     *
     * @param file 上传的 MultipartFile
     * @return 字数
     */
    public static int countWordsIgnoringImages(MultipartFile file) throws IOException {
        if (file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }

        String filename = file.getOriginalFilename();
        if (filename == null) {
            throw new IllegalArgumentException("文件名不能为空");
        }

        String lowerName = filename.toLowerCase();

        if (lowerName.endsWith(".txt")) {
            return countInTxt(file);
        } else if (lowerName.endsWith(".docx")) {
            return countInDocx(file);
        } else if (lowerName.endsWith(".pdf")) {
            return countInPdf(file);
        } else if (lowerName.endsWith(".pptx")) {
            return countInPptx(file);
        } else if (lowerName.endsWith(".xlsx")) {
            return countInXlsx(file);
        } else {
            throw new UnsupportedOperationException("不支持的文件格式: " + filename);
        }
    }

    // ==================== TXT ====================
    private static int countInTxt(MultipartFile file) throws IOException {
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(file.getInputStream(), StandardCharsets.UTF_8))) {
            int count = 0;
            int ch;
            while ((ch = reader.read()) != -1) {
                char c = (char) ch;
                if (CHINESE_PATTERN.matcher(String.valueOf(c)).matches() ||
                        LETTER_PATTERN.matcher(String.valueOf(c)).matches()) {
                    count++;
                }
            }
            return count;
        }
    }

    // ==================== DOCX ====================
    private static int countInDocx(MultipartFile file) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(file.getInputStream())) {
            int count = 0;
            List<XWPFParagraph> paragraphs = doc.getParagraphs();

            for (XWPFParagraph paragraph : paragraphs) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    for (XWPFRun run : runs) {
                        // 忽略图片（POI 中可通过判断是否有图片来跳过，但 run 本身不直接暴露图片）
                        // 实际上，XWPFRun.getImage() 不存在，图片是通过 Drawing 对象管理的
                        // 所以我们默认提取 run.getText()，它不会包含图片内容
                        String text = run.getText(0); // 提取文本
                        if (text != null) {
                            count += countLettersAndChinese(text);
                        }
                    }
                }
            }
            return count;
        }
    }

    // ==================== PDF ====================
    private static int countInPdf(MultipartFile file) throws IOException {
        try (PDDocument document = Loader.loadPDF(new RandomAccessReadBuffer(file.getInputStream()))) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);
            return countLettersAndChinese(text);
        }
    }

    // ==================== PPTX ====================
    private static int countInPptx(MultipartFile file) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow(file.getInputStream())) {
            int count = 0;
            List<XSLFSlide> slides = ppt.getSlides();

            for (XSLFSlide slide : slides) {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape) shape;
                        String text = textShape.getText();
                        if (text != null) {
                            count += countLettersAndChinese(text);
                        }
                    }
                    // 其他形状（如图片、图表）被忽略
                }
            }
            return count;
        }
    }

    // ==================== XLSX ====================
    private static int countInXlsx(MultipartFile file) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream())) {
            int count = 0;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                for (org.apache.poi.ss.usermodel.Row row : sheet) {
                    for (org.apache.poi.ss.usermodel.Cell cell : row) {
                        String text = "";
                        switch (cell.getCellType()) {
                            case STRING:
                                text = cell.getStringCellValue();
                                break;
                            case FORMULA:
                                // 可选：处理公式结果
                                break;
                            case NUMERIC:
                                text = String.valueOf(cell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                text = String.valueOf(cell.getBooleanCellValue());
                                break;
                            default:
                                break;
                        }
                        count += countLettersAndChinese(text);
                    }
                }
            }
            return count;
        }
    }

    // ==================== 工具方法 ====================
    /**
     * 统计字符串中的中文字 + 英文字母数量
     */
    private static int countLettersAndChinese(String text) {
        int count = 0;
        for (char c : text.toCharArray()) {
            if (CHINESE_PATTERN.matcher(String.valueOf(c)).matches() ||
                    LETTER_PATTERN.matcher(String.valueOf(c)).matches()) {
                count++;
            }
        }
        return count;
    }
}
