package com.zlz.split.file.splitfile.controller;

import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URLDecoder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

@RestController
@Slf4j
@RequestMapping("/file")
public class Md2WordController {

    @PostMapping("/md-convert-word")
    public byte[] convertMarkdownToDocxV2(@RequestBody Map<String, String> request)
            throws IOException, InterruptedException {
        String markdownContent = request.get("markdown");
        String fileName = request.getOrDefault("fileName", "document");
        markdownContent = markdownContent.replaceAll("/foundryFile/files", "http://aifoundry.unisoc.com:8099/files");
        log.info("----convert-word--md文件：{} 转为word文件", fileName);

        Path tempMd = Files.createTempFile("upload_", ".md");
        Path tempDocx = Files.createTempFile("output_", ".docx");

        try {
            Files.writeString(tempMd, markdownContent);
            // 构建 Pandoc 命令（支持数学公式）
            ProcessBuilder pb = new ProcessBuilder(
                    "pandoc",
                    "--mathml",               // 或 --mathjax（需网络），推荐 --mathml 兼容性好
                    "--wrap=none",            // 避免自动换行破坏公式
                    tempMd.toString(),
                    "-o", tempDocx.toString()
            );

            // 设置工作目录（可选，确保图片路径正确）
            pb.directory(new File(System.getProperty("user.dir")));

            Process process = pb.start();

            // 读取错误输出（便于调试）
            try (BufferedReader errorReader = new BufferedReader(
                    new InputStreamReader(process.getErrorStream()))) {
                String line;
                while ((line = errorReader.readLine()) != null) {
                    log.error("[Pandoc Error] " + line);
                }
            }

            int exitCode = process.waitFor();
            if (exitCode == 0) {
                log.info("✅ 转换成功: {}", tempDocx.toString());
                return Files.readAllBytes(tempDocx);
            } else {
                throw new RuntimeException("Pandoc 转换失败，退出码: " + exitCode);
            }
        } finally {
            Files.deleteIfExists(tempMd);
            Files.deleteIfExists(tempDocx);
        }
    }

    /**
     * md格式直接转换为doc，可供前端直接下载
     */
    @PostMapping("/direct/md-convert-word")
    public void convertMarkdownToDocx(HttpServletResponse response, @RequestBody Map<String, String> request)
            throws IOException, InterruptedException {
        String decodeFileName = URLDecoder.decode(request.get("fileName"), "UTF-8");
        String markdownContent = request.get("markdown");
        String fileName = request.getOrDefault("fileName", "document");
        markdownContent = markdownContent.replaceAll("/foundryFile/files", "http://aifoundry.unisoc.com:8099/files");
        log.info("----convert-word--md文件：{} 直接转为word文件", fileName);

        Path tempMd = Files.createTempFile("upload_", ".md");
        Path tempDocx = Files.createTempFile("output_", ".docx");

        try {
            Files.writeString(tempMd, markdownContent);
        // 构建 Pandoc 命令（支持数学公式）
        ProcessBuilder pb = new ProcessBuilder(
                "pandoc",
                "--mathml",               // 或 --mathjax（需网络），推荐 --mathml 兼容性好
                "--wrap=none",            // 避免自动换行破坏公式
                tempMd.toString(),
                "-o", tempDocx.toString()
        );

        // 设置工作目录（可选，确保图片路径正确）
        pb.directory(new File(System.getProperty("user.dir")));

        Process process = pb.start();

        // 读取错误输出（便于调试）
        try (BufferedReader errorReader = new BufferedReader(
                new InputStreamReader(process.getErrorStream()))) {
            String line;
            while ((line = errorReader.readLine()) != null) {
                log.error("[Pandoc Error] " + line);
            }
        }

        int exitCode = process.waitFor();
        if (exitCode == 0) {
            log.info("✅ 转换成功: {}", tempDocx.toString());
            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.setHeader("Content-Disposition", "attachment; filename=\"" +
                    java.net.URLEncoder.encode(decodeFileName.replace(".md", ".docx"), "UTF-8") + "\"");

            Files.copy(tempDocx, response.getOutputStream());
            response.getOutputStream().flush();
        } else {
            throw new RuntimeException("Pandoc 转换失败，退出码: " + exitCode);
        }
        } finally {
            Files.deleteIfExists(tempMd);
            Files.deleteIfExists(tempDocx);
        }
    }

}
