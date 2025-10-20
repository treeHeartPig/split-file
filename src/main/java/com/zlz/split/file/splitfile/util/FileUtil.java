package com.zlz.split.file.splitfile.util;

import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jodconverter.core.DocumentConverter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
@Service
public class FileUtil {

    @Value("${libreoffice.tmp-path}")
    private String libreofficeTmpPath;

    @Autowired
    @Qualifier("defaultConverter")
    private DocumentConverter converter;

    public static File getFontFile(String path) throws IOException {
        Resource resource = new ClassPathResource( path);
        File tmpFile = File.createTempFile("tmp"+System.currentTimeMillis(), ".ttf");
        tmpFile.deleteOnExit();
        try(InputStream inputStream = resource.getInputStream()){
            Path tmpPath = tmpFile.toPath();
            Files.copy(inputStream,tmpPath, StandardCopyOption.REPLACE_EXISTING);
            return  tmpFile;
        }
    }

    public String processLargeDocument(File file) throws Exception {

        // 2. 切分大文档
        List<File> splitFiles = splitDocument(file);

        // 3. 转换为PDF
        List<File> pdfFiles = new ArrayList<>();
        for (File docFile : splitFiles) {
            File pdfFile = convertToPdf(docFile);
            pdfFiles.add(pdfFile);
        }

        // 4. 合并PDF
        File mergedPdf = mergePdfFiles(pdfFiles);

        // 5. 清理临时文件
        cleanTempFiles(splitFiles, pdfFiles);

        return mergedPdf.getAbsolutePath();
    }

    private List<File> splitDocument(File originalFile) throws Exception {
        // 使用Apache POI读取Word文档并按页或章节切分
        // 这里简化为示例，实际应根据需求实现切分逻辑
        List<File> splitFiles = new ArrayList<>();

        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(originalFile))) {
            int pageCount = doc.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();

            // 假设每10页切分为一个文件
            int pagesPerFile = 10;
            int fileCount = (int) Math.ceil((double) pageCount / pagesPerFile);

            for (int i = 0; i < fileCount; i++) {
                XWPFDocument newDoc = new XWPFDocument();
                // 这里需要实现实际的内容复制逻辑
                // 简单示例：创建一个包含部分内容的新文档
                String splitFileName = originalFile.getName().replace(".docx", "") +
                        "_part" + (i + 1) + ".docx";
                File splitFile = new File(libreofficeTmpPath + File.separator + splitFileName);

                try (FileOutputStream out = new FileOutputStream(splitFile)) {
                    newDoc.write(out);
                }

                splitFiles.add(splitFile);
            }
        }

        return splitFiles;
    }

    private File convertToPdf(File docFile) throws Exception {
        String pdfPath = docFile.getAbsolutePath().replace(".docx", ".pdf");
        File pdfFile = new File(pdfPath);
        converter.convert(docFile).to(pdfFile).execute();
        return pdfFile;
    }

    private File mergePdfFiles(List<File> pdfFiles) throws IOException {
        PDFMergerUtility merger = new PDFMergerUtility();
        String mergedFileName = "merged_" + System.currentTimeMillis() + ".pdf";
        File mergedFile = new File(libreofficeTmpPath + File.separator + mergedFileName);

        for (File pdfFile : pdfFiles) {
            merger.addSource(pdfFile);
        }

        merger.setDestinationFileName(mergedFile.getAbsolutePath());
        merger.mergeDocuments(MemoryUsageSetting.setupMainMemoryOnly());

        return mergedFile;
    }

    private void cleanTempFiles(List<File> splitFiles, List<File> pdfFiles) {
        // 删除临时文件
        for (File file : splitFiles) {
            file.delete();
        }
        for (File file : pdfFiles) {
            file.delete();
        }
    }
}
