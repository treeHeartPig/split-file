package com.zlz.split.file.splitfile.controller;

import com.zlz.split.file.splitfile.service.ExcelSplitterService;
import com.zlz.split.file.splitfile.service.FileProcessor;
import com.zlz.split.file.splitfile.util.MinioUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.util.StringUtil;
import org.jetbrains.annotations.NotNull;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeManager;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Random;

@RestController
@Slf4j
@RequestMapping("/file")
public class FileSplitController {
    @Autowired
    private ExcelSplitterService excelSplitterService;
    @Autowired
    private FileProcessor fileProcessor;

    @Autowired
    private OfficeManager officeManager;
    @Autowired
    private MinioUtil minioUtil;
    @Autowired
    private DocumentConverter defaultConverter;

    @RequestMapping(value = {"/split","/upload"}, method = RequestMethod.POST)
    public Map<String,Object> uploadFile(MultipartFile file,String sn) throws Exception{
        String decodeFileName = URLDecoder.decode(file.getOriginalFilename(), "utf-8");
        String baseName = checkBaseNameLength(decodeFileName);
        File tmpFile = null;
        File pdfFile = null;
        if(StringUtil.isBlank(sn)) sn = "default";
        try{
            String extension = FilenameUtils.getExtension(file.getOriginalFilename());
            List<String> excelTypes = new ArrayList<>();
            excelTypes.add("xlsx");
            excelTypes.add("xls");
            if(excelTypes.contains(extension)){
                return excelSplitterService.splitExcelBySheet(file,sn);
            }
            tmpFile  = File.createTempFile(baseName, "." + extension);
            file.transferTo(tmpFile);

            pdfFile = fileProcessor.convertToPdf(tmpFile, extension);
            return fileProcessor.processFile(pdfFile,baseName,sn);
        }finally {
            if(tmpFile != null){
                tmpFile.delete();
            }
            if(pdfFile != null){
                pdfFile.delete();
            }
        }
    }

    @RequestMapping(value = {"/convertToPdf"}, method = RequestMethod.POST)
    public String convertToPdf(@RequestParam("file") MultipartFile file, @RequestParam("sn") String sn) throws Exception {
        log.info("开始转换文件：{}为PDF格式", file.getOriginalFilename());
        String decodeFileName = URLDecoder.decode(file.getOriginalFilename(), "UTF-8");
        String baseName = checkBaseNameLength(decodeFileName);
        if (StringUtil.isBlank(sn)) sn = "default";

        File tmpFile = null;
        File pdfFile = null;
        try {
            // 1. 创建临时文件
            String extension = FilenameUtils.getExtension(file.getOriginalFilename());
            tmpFile = File.createTempFile(baseName, "." + extension);
            file.transferTo(tmpFile);
            // 2. 使用 JODConverter 转换
            pdfFile = fileProcessor.convertToPdf(tmpFile, extension);

            // 3. 检查 PDF 是否生成成功
            if (!pdfFile.exists() || pdfFile.length() == 0) {
                throw new Exception("PDF 转换失败，输出文件为空");
            }

            // 4. 上传到 MinIO
            String pdfPath = minioUtil.streamUploadToMinio(pdfFile, baseName + ".pdf", "application/pdf", sn);
            String wholePath = minioUtil.getFileUrl(pdfPath);

            log.info("完成 转换文件：{}为PDF格式",file.getOriginalFilename());
            return wholePath;

        } finally {
            if (tmpFile != null && tmpFile.exists()) tmpFile.delete();
            if (pdfFile != null && pdfFile.exists()) pdfFile.delete();
        }
    }


    @NotNull
    private static String checkBaseNameLength(String decodeFileName) {
        String baseName = FilenameUtils.getBaseName(decodeFileName);
        baseName=baseName.replaceAll("\\s+","-");
        if (baseName.length() < 3) {
            StringBuilder sb = new StringBuilder(baseName);
            Random random = new Random();
            String chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            while (sb.length() < 3) {
                int index = random.nextInt(chars.length());
                sb.append(chars.charAt(index));
            }
            baseName = sb.toString();
        }
        return baseName;
    }
}
