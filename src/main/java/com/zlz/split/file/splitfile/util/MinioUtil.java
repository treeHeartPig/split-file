package com.zlz.split.file.splitfile.util;


import com.zlz.split.file.splitfile.config.MinioConfig;
import io.minio.MinioClient;
import io.minio.PutObjectArgs;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.util.UUID;


@Service
@Slf4j
public class MinioUtil {
    @Autowired
    private MinioClient minioClient;
    @Autowired
    private MinioConfig minioConfig;
    private static final String DEFAULT_READ_ASSISTANT_DIR= "read-assistant";

    public String getFileUrl(String fileName){
        try{
            return minioConfig.getUrl() +"/"+ minioConfig.getBucketName() +"/"+ fileName;
        }catch (Exception e){
            log.error("--minioUtil#getFileUrl返回异常",e);
            throw e;
        }
    }
    public String uploadToMinio(byte[] data, String objectName, String contentType) throws Exception {
        String path = DEFAULT_READ_ASSISTANT_DIR + "/" + UUID.randomUUID() + "/" +objectName;
        ByteArrayInputStream bais = new ByteArrayInputStream(data);
        minioClient.putObject(
                PutObjectArgs.builder()
                        .bucket(minioConfig.getBucketName())
                        .object(path)
                        .stream(bais, data.length, -1)
                        .contentType(contentType)
                        .build());
        return path;
    }
}
