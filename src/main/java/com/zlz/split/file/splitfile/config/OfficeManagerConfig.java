package com.zlz.split.file.splitfile.config;

import com.zlz.split.file.splitfile.util.OsUtil;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.office.LocalOfficeContext;
import org.jodconverter.local.office.LocalOfficeManager;
import org.jodconverter.local.process.ProcessManager;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class OfficeManagerConfig {
    // 配置进程管理器（启动 4 个 LibreOffice 进程）LocalOfficeProcessManager(4)
//    ProcessManager processManager = new ProcessPoolOfficeProcessManager(4);

    @Bean
    public OfficeManager getOfficeManager() throws OfficeException {
        OsUtil.OS os = OsUtil.getOS();
        OfficeManager officeManager;
        if (os == OsUtil.OS.WINDOWS) {
            officeManager = LocalOfficeManager.builder()
                    .officeHome("C:\\Program Files\\LibreOffice")  // LibreOffice 路径
//                    .processManager()                // 进程管理器
//                    .portNumbers(2002, 2003)          // 每个进程的独立端口
                    .maxTasksPerProcess(50)                      // 进程重启阈值
                    .taskQueueTimeout(30000L)                     // 任务队列超时（毫秒）
                    .build();
        }else{
            officeManager = LocalOfficeManager.builder().officeHome("/usr/lib/libreoffice")
//                    .processManager(processManager)                // 进程管理器
//                    .portNumbers(2002, 2003)          // 每个进程的独立端口
                    .maxTasksPerProcess(50)                      // 进程重启阈值s
                    .taskQueueTimeout(30000L)                     // 任务队列超时（毫秒）
                    .build();
        }
        return officeManager;
    }
}
