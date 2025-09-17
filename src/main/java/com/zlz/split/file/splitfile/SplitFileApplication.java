package com.zlz.split.file.splitfile;

import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import lombok.extern.slf4j.Slf4j;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.office.LocalOfficeManager;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

import javax.annotation.PreDestroy;

@Slf4j
@SpringBootApplication
public class SplitFileApplication {
 private static OfficeManager manager;
    public static void main(String[] args) {
        ConfigurableApplicationContext context = SpringApplication.run(SplitFileApplication.class, args);
        manager = context.getBean(OfficeManager.class);
    }

    @PreDestroy
    public void destroy() {
        log.info("--尝试关闭officeManager……");
        if(manager != null && manager.isRunning()){
            try {
                manager.stop();
                log.info("--完成关闭officeManager");
            } catch (OfficeException e) {
                log.error("关闭officeManager失败",e);
                throw new RuntimeException(e);
            }
        }else {
            log.info("--officeManager已关闭");
        }
    }
}
