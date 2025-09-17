package com.zlz.split.file.splitfile.controller;

import com.zlz.split.file.splitfile.service.ExcelSplitterService;
import com.zlz.split.file.splitfile.service.FileProcessor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FilenameUtils;
import org.jodconverter.core.office.OfficeManager;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

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
    @PostMapping("/upload")
    public Map<String,Map<Integer,String>> uploadFile(MultipartFile file) throws Exception{
        String baseName = FilenameUtils.getBaseName(file.getOriginalFilename());
        File tmpFile = null;
        File pdfFile = null;
        try{
            String extension = FilenameUtils.getExtension(file.getOriginalFilename());
            List<String> excelTypes = new ArrayList<>();
            excelTypes.add("xlsx");
            excelTypes.add("xls");
            if(excelTypes.contains(extension)){
                return excelSplitterService.splitExcelBySheet(file);
            }
            tmpFile  = File.createTempFile(baseName, "." + extension);
            file.transferTo(tmpFile);

            pdfFile = fileProcessor.convertToPdf(tmpFile, extension);
            return fileProcessor.processFile(pdfFile,baseName);
        }finally {
            if(tmpFile != null){
                tmpFile.delete();
            }
            if(pdfFile != null){
                pdfFile.delete();
            }
        }
    }
}
