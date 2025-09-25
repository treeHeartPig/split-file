package com.zlz.split.file.splitfile.config;

import java.awt.*;
import java.io.File;

public class FontLoaderConfig {
    public static void registerChineseFonts(String fontPath){
        File fontDir = new File(fontPath);
        if(!fontDir.exists() || !fontDir.isDirectory()){
            System.out.println("--字体目录不存在");
            return;
        }
        for(File fontFile : fontDir.listFiles((dir,name) -> name.toLowerCase().endsWith(".ttf")
                || name.toLowerCase().endsWith(".ttc"))){
            try{
                Font font = Font.createFont(Font.TRUETYPE_FONT, fontFile);
                GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
                ge.registerFont(font);
                System.out.println("--加载字体成功：" + fontFile.getName());
            }catch (Exception e){
                System.out.println("--加载字体失败：" + fontFile.getName());
            }
        }
    }
}
