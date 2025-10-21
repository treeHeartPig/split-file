package com.zlz.split.file.splitfile.controller;

import com.zlz.split.file.splitfile.service.FileCounterService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@RestController
@Slf4j
@RequestMapping("/file")
public class FileCountController {
    @Autowired
    private FileCounterService fileCounterService;
    @PostMapping("/count-words")
    public ResponseEntity<Integer> countWords(@RequestParam("file") MultipartFile file) {
        try {
            int wordCount = fileCounterService.countWordsIgnoringImages(file);
            System.out.println("---文件:"+file.getOriginalFilename()+"--字数共："+wordCount);
            return ResponseEntity.ok(wordCount);
        } catch (IOException e) {
            log.error("--------Error counting words-----", e);
            return ResponseEntity.badRequest().body(0);
        } catch (UnsupportedOperationException e) {
            log.error("--------统计字数接口UnsupportedOperationException----", e);
            return ResponseEntity.badRequest().body(-1);
        }
    }

    @PostMapping("/count-pages")
    public ResponseEntity<Integer> countPages(@RequestParam("file") MultipartFile file) {
        try(InputStream is = file.getInputStream()) {
            int pageCount = fileCounterService.countPages(is,file.getOriginalFilename());
            log.info("----文件:{}共{}页",file.getOriginalFilename(),pageCount);
            return ResponseEntity.ok(pageCount);
        } catch (Exception e) {
            log.error("--------统计页数接口UnsupportedOperationException----", e);
            return ResponseEntity.badRequest().body(-1);
        }
    }
}
