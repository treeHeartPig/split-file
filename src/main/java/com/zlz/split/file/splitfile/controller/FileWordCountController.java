package com.zlz.split.file.splitfile.controller;

import com.zlz.split.file.splitfile.service.FileWordCounter;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@Slf4j
@RequestMapping("/file")
public class FileWordCountController {
    @PostMapping("/count-words")
    public ResponseEntity<Integer> countWords(@RequestParam("file") MultipartFile file) {
        try {
            int wordCount = FileWordCounter.countWordsIgnoringImages(file);
            return ResponseEntity.ok(wordCount);
        } catch (IOException e) {
            log.error("--------Error counting words-----", e);
            return ResponseEntity.badRequest().body(0);
        } catch (UnsupportedOperationException e) {
            log.error("--------统计字数接口UnsupportedOperationException----", e);
            return ResponseEntity.badRequest().body(-1);
        }
    }
}
