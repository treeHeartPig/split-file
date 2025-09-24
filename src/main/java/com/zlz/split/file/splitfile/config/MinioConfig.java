package com.zlz.split.file.splitfile.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@ConfigurationProperties(prefix = "minio")
@Component
@Data
public class MinioConfig {
    private String url;
    private String accessKey;
    private String secretKey;
    private String bucketName;
}
