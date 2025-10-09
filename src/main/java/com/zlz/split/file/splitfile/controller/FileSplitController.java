package com.zlz.split.file.splitfile.controller;

import com.zlz.split.file.splitfile.service.ExcelSplitterService;
import com.zlz.split.file.splitfile.service.FileProcessor;
import com.zlz.split.file.splitfile.util.MinioUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.util.StringUtil;
import org.jetbrains.annotations.NotNull;
import org.jodconverter.core.office.OfficeManager;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
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
    public String convertToPdf(MultipartFile file,String sn) throws Exception{
        String decodeFileName = URLDecoder.decode(file.getOriginalFilename(), "utf-8");
        String baseName = checkBaseNameLength(decodeFileName);
        File tmpFile = null;
        File pdfFile = null;
        if(StringUtil.isBlank(sn)) sn = "default";
        try{
            String extension = FilenameUtils.getExtension(file.getOriginalFilename());
            tmpFile  = File.createTempFile(baseName, "." + extension);
            file.transferTo(tmpFile);

            pdfFile = fileProcessor.convertToPdf(tmpFile, extension);
            String pdfPath = minioUtil.streamUploadToMinio(pdfFile, baseName+".pdf", "application/pdf", sn);
            String wholePath = minioUtil.getFileUrl(pdfPath);
            System.out.println(String.format("-文件:%s 转为pdf格式后，路径为:%s",baseName,wholePath));
            return wholePath;
        }finally {
            if(tmpFile != null){
                tmpFile.delete();
            }
            if(pdfFile != null){
                pdfFile.delete();
            }
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
