package com.zlz.split.file.splitfile.config;

import io.minio.MinioClient;
import lombok.extern.slf4j.Slf4j;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;
import org.springframework.beans.factory.DisposableBean;
import org.springframework.beans.factory.InitializingBean;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import static com.zlz.split.file.splitfile.constants.Constant.*;
@Slf4j
@Configuration
public class AppConfig implements InitializingBean, DisposableBean {
    @Autowired
    private OfficeManager officeManager;
    @Autowired
    private MinioConfig minioConfig;
    private DocumentConverter converter;
    @Override
    public void afterPropertiesSet() throws Exception {
        officeManager.start();
    }

    @Bean
    public DocumentConverter getDocumentConverter() throws OfficeException {
        return LocalConverter.builder().officeManager(officeManager).build();
    }

    @Bean
    public MinioClient getMinioClient(){
        return MinioClient.builder()
                .endpoint(minioConfig.getUrl())
                .credentials(minioConfig.getAccessKey(), minioConfig.getSecretKey())
                .build();
    }

    @Override
    public void destroy() throws Exception {
        if(officeManager != null && officeManager.isRunning()){
            new Thread(() -> {
                try {
                    officeManager.stop();
                }catch (OfficeException e){
                    log.error("关闭officeManager失败",e);
                }
            }).start();
        }
        if(converter != null){
            new Thread(() -> {
                if(converter instanceof AutoCloseable){
                    try {
                        ((AutoCloseable)converter).close();
                    }catch (Exception e){
                        log.error("关闭converter失败",e);
                    }
                }
            }).start();
        }
    }
}
