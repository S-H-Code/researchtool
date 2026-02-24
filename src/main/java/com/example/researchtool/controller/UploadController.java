package com.example.researchtool.controller;

import com.example.researchtool.service.ExtractionService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;

@Controller
public class UploadController {

    @Autowired
    private ExtractionService extractionService;

    @GetMapping("/")
    public String home() {
        return "index";
    }
@PostMapping("/upload")
public ResponseEntity<InputStreamResource> uploadFile(
        @RequestParam("file") MultipartFile file) throws Exception {

    ByteArrayInputStream excelFile = extractionService.processPdf(file);

    HttpHeaders headers = new HttpHeaders();
    headers.add("Content-Disposition", "attachment; filename=output.xlsx");

    return ResponseEntity
            .ok()
            .headers(headers)
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .body(new InputStreamResource(excelFile));
}
}