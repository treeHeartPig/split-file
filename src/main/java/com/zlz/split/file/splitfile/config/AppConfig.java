package com.zlz.split.file.splitfile.config;

import com.zlz.split.file.splitfile.util.OsUtil;
import io.minio.MinioClient;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import static com.zlz.split.file.splitfile.constants.Constant.*;

@Configuration
public class AppConfig {
    @Bean
    public OfficeManager getOfficeManager() throws OfficeException {
        OsUtil.OS os = OsUtil.getOS();
        OfficeManager officeManager;
        if (os == OsUtil.OS.WINDOWS) {
            officeManager = LocalOfficeManager.builder().officeHome("C:\\Program Files\\LibreOffice").build();
        }else{
            officeManager = LocalOfficeManager.builder().officeHome("/Applications/LibreOffice.app/Contents").build();
        }
        officeManager.start();
        return officeManager;
    }

    @Bean
    public MinioClient getMinioClient(){
        return MinioClient.builder()
                .endpoint(MINIO_URL)
                .credentials(MINIO_ADMIN, MINIO_PWD)
                .build();
    }

    @Bean
    public DocumentConverter getDocumentConverter() throws OfficeException {
        return LocalConverter.make(getOfficeManager());
    }
}
