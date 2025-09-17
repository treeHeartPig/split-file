package com.zlz.split.file.splitfile.util;

public class OsUtil {
    public  enum OS {
        WINDOWS,
        MAC,
        LINUX,
        OTHER
    }
    public static OS getOS() {
        String os = System.getProperty("os.name").toLowerCase();
        if (os.contains("win")) {
            return OS.WINDOWS;
        } else if (os.contains("mac")) {
            return OS.MAC;
        } else if (os.contains("nix") || os.contains("nux") || os.contains("aix")) {
            return OS.LINUX;
        } else {
            return OS.OTHER;
        }
    }
}

