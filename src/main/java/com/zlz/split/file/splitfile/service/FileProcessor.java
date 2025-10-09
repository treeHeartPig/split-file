package com.zlz.split.file.splitfile.service;

import com.alibaba.fastjson.JSONObject;
import com.zlz.split.file.splitfile.util.MinioUtil;
import io.minio.MinioClient;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FilenameUtils;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.util.StringUtil;
import org.jodconverter.core.DocumentConverter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.net.URLDecoder;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.zlz.split.file.splitfile.util.FileUtil.getFontFile;

@Slf4j
@Service
public class FileProcessor {
    @Autowired
    private MinioClient minioClient;
    @Autowired
    private DocumentConverter converter;
    @Autowired
    private MinioUtil minioUtil;

    public Map<String,Object> processFile(File pdfFile, String baseName,String sn) throws Exception {
        Map<String,Object> result = new HashMap<>();
        // 使用内存优化设置加载PDF
        try (PDDocument document = PDDocument.load(pdfFile,
                org.apache.pdfbox.io.MemoryUsageSetting.setupTempFileOnly())) {
            List<String> pageObjectNames = new ArrayList<>();
            int totalPages = document.getNumberOfPages();
            result.put("totalPages",totalPages);

            Map<Integer,String> pages = new HashMap<>();
            Map<Integer,String> thumbnails = new HashMap<>();
            // 分批处理页面，避免同时占用过多内存
            for (int i = 0; i < totalPages; i++) {
                try {
                    // Step 1: 按页切分并上传 PDF 子页面
                    String pagePdfName = String.format("%s_page_%d.pdf", baseName, i + 1);
                    byte[] pagePdfBytes = extractPageAsPdf(document, i);
                    String path = minioUtil.uploadToMinio(pagePdfBytes, pagePdfName, "application/pdf",sn);
                    pages.put(i,minioUtil.getFileUrl(path));
                    pageObjectNames.add(pagePdfName);

                    // Step 2: 生成缩略图
                    // 使用较低DPI减少内存占用
                    BufferedImage img = renderPageImage(document, i, 72);
                    if (img != null) {
                        byte[] thumbData = convertImageToBytes(img, "jpg");
                        if (thumbData != null) {
                            String thumbName = String.format("%s_thumb_%d.jpg", baseName, i + 1);
                            String path2 = minioUtil.uploadToMinio(thumbData, thumbName, "image/jpeg",sn);
                            thumbnails.put(i,minioUtil.getFileUrl(path2));
                        }
                    }
                    // 及时清理页面相关资源
                    clearPageResources(img);
                    result.put("pages",pages);
                    result.put("thumbnails",thumbnails);
                } catch (OutOfMemoryError e) {
                    System.err.printf("Out of memory when processing page %d, trying with lower quality%n", i + 1);
                    // 尝试使用更低质量处理
                    processPageWithLowMemory(document, i, baseName, pageObjectNames,sn);
                } catch (Exception e) {
                    System.err.printf("Error processing page %d: %s%n", i + 1, e.getMessage());
                    // 继续处理下一页而不是中断整个过程
                }
            }
        }
        System.out.println("------split-result:"+ JSONObject.toJSONString(result));
        return result;
    }

    // 优化的页面渲染方法
    private BufferedImage renderPageImage(PDDocument document, int pageIndex, int dpi) {
        try {
            PDFRenderer renderer = new PDFRenderer(document);
            // 使用特定的渲染参数减少内存使用
            return renderer.renderImageWithDPI(pageIndex, dpi,
                    org.apache.pdfbox.rendering.ImageType.RGB);
        } catch (Exception e) {
            System.err.printf("Error rendering page %d: %s%n", pageIndex + 1, e.getMessage());
            return null;
        }
    }

    // 内存不足时的降级处理
    private void processPageWithLowMemory(PDDocument document, int pageIndex,
                                          String baseName, List<String> pageObjectNames,String sn) throws Exception {
        try {
            // 使用更低的DPI重新渲染
            BufferedImage img = renderPageImage(document, pageIndex, 36); // 降低到36 DPI
            if (img != null) {
                // 重新处理该页面
                String pagePdfName = String.format("%s_page_%d.pdf", baseName, pageIndex + 1);
                byte[] pagePdfBytes = extractPageAsPdf(document, pageIndex);
                String pagePdfPath = minioUtil.uploadToMinio(pagePdfBytes, pagePdfName, "application/pdf",sn);
                pageObjectNames.add(pagePdfPath);

                byte[] thumbData = convertImageToBytes(img, "jpg");
                if (thumbData != null) {
                    String thumbName = String.format("%s_thumb_%d.jpg", baseName, pageIndex + 1);
                    String thumbPath = minioUtil.uploadToMinio(thumbData, thumbName, "image/jpeg",sn);
                    System.out.printf("Uploaded Page %d (low quality): PDF=%s, Thumb=%s%n",
                            pageIndex+1, pagePdfPath, thumbPath);
                }

                clearPageResources(img);
            }
        } catch (Exception e) {
            System.err.printf("Failed to process page %d even with low quality: %s%n",
                    pageIndex + 1, e.getMessage());
            throw e;
        }
    }

    // 优化的图像转换方法
    private byte[] convertImageToBytes(BufferedImage img, String format) {
        try (ByteArrayOutputStream imgOut = new ByteArrayOutputStream()) {
            ImageIO.write(img, format, imgOut);
            return imgOut.toByteArray();
        } catch (Exception e) {
            System.err.printf("Error converting image to bytes: %s%n", e.getMessage());
            return null;
        }
    }

    // 清理页面资源
    private void clearPageResources(BufferedImage img) {
        if (img != null) {
            img.flush(); // 释放图像资源
        }
        // 建议进行垃圾回收提示（谨慎使用）
        // System.gc();
    }


    public File convertToPdf(File inputFile, String ext) throws Exception {
        String decodeFileName = URLDecoder.decode(inputFile.getName(), "utf-8");
        String fileName = FilenameUtils.getBaseName(decodeFileName);
        File outputPdf = new File(inputFile.getParent(), "temp_"+fileName+".pdf");

        switch (ext) {
            case "txt":
                // TXT 转 PDF（简单包装文本）
                txtToPdf(inputFile, outputPdf);
                break;
            case "docx":
            case "doc":
            case "xlsx":
            case "xls":
            case "pptx":
                converter.convert(inputFile).to(outputPdf).execute();
                break;
            case "pdf":
                Files.copy(inputFile.toPath(), outputPdf.toPath());
                return inputFile; // 原始就是 PDF
            default:
                throw new IllegalArgumentException("Unsupported file type: " + ext);
        }

        return outputPdf;
    }

    private byte[] extractPageAsPdf(PDDocument fullDoc, int pageIndex) throws IOException {
        try (PDDocument singlePageDoc = new PDDocument()) {
            singlePageDoc.addPage(fullDoc.getPage(pageIndex));
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            singlePageDoc.save(baos);
            return baos.toByteArray();
        }
    }

    private void txtToPdf(File txtFile, File pdfFile) throws IOException {
        try (BufferedReader reader = Files.newBufferedReader(txtFile.toPath());
             PDDocument doc = new PDDocument()) {
            // 简单文本转PDF（仅示例，可扩展排版）
            PDPage page = new PDPage();
            doc.addPage(page);
            // 从本地字体文件加载字体
            PDType0Font font = PDType0Font.load(doc, getFontFile("fonts/msyh.ttf"));

            try (PDPageContentStream content = new PDPageContentStream(doc, page)) {
                content.beginText();
                content.setFont(font, 12);
                content.newLineAtOffset(50, 700);
                String line;
                float y = 700;
                while ((line = reader.readLine()) != null && y > 50) {
                    content.showText(line);
                    content.newLineAtOffset(0, -15);
                    y -= 15;
                }
                content.endText();
            }
            doc.save(pdfFile);
        }
    }

}