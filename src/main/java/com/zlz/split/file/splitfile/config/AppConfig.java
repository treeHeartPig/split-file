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
import org.springframework.context.annotation.Lazy;

import java.util.HashMap;
import java.util.Map;

@Slf4j
@Configuration
public class AppConfig implements InitializingBean, DisposableBean {
    @Autowired
    private OfficeManager officeManager;
    @Autowired
    private MinioConfig minioConfig;
    private DocumentConverter converter;
    private DocumentConverter excelConverter;
    @Override
    public void afterPropertiesSet() throws Exception {
        officeManager.start();
    }

    @Bean("defaultConverter")
    @Lazy
    public DocumentConverter getDocumentConverter() throws OfficeException {
        return LocalConverter.builder().officeManager(officeManager).build();
    }


    @Bean("excelConverter")
    @Lazy
    public DocumentConverter getCustomDocumentConverter() throws OfficeException {
        Map<String, Object> defaultLoadProps = new HashMap<>();
        defaultLoadProps.put("Hidden", true);

        Map<String, Object> defaultStoreProps = new HashMap<>();
        defaultStoreProps.put("FitToPages", true);
        defaultStoreProps.put("FitToPagesX", 1);
        defaultStoreProps.put("FitToPagesY", 0);
        return LocalConverter.builder()
                .officeManager(officeManager)
                .loadProperties(defaultLoadProps)
                .storeProperties(defaultStoreProps)
                .build();
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
        if(excelConverter != null){
            new Thread(() -> {
                if(excelConverter instanceof AutoCloseable){
                    try {
                        ((AutoCloseable)excelConverter).close();
                    }catch (Exception e){
                        log.error("关闭converter失败",e);
                    }
                }
            }).start();
        }
    }
}
