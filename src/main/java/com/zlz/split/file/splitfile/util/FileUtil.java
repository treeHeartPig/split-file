package com.zlz.split.file.splitfile.util;

import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

public class FileUtil {
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
}
