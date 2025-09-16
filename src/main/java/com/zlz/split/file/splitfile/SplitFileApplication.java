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

@Slf4j
@SpringBootApplication
public class SplitFileApplication {

    public static void main(String[] args) {
        ConfigurableApplicationContext context = SpringApplication.run(SplitFileApplication.class, args);
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            log.info("--开始关闭office");
                OfficeManager manager = context.getBean(OfficeManager.class);
                if(manager.isRunning()){
                    try {
                        manager.stop();
                        log.info("--完成关闭office");
                    } catch (OfficeException e) {
                        log.error("关闭office失败",e);
                        throw new RuntimeException(e);
                    }
                }
        }));
    }
}
