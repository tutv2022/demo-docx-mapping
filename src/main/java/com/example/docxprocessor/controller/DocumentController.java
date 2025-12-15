package com.example.docxprocessor.controller;

import com.example.docxprocessor.model.TemplateData;
import com.example.docxprocessor.service.DocxTemplateService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.UUID;

@RestController
@RequestMapping("/api/documents")
public class DocumentController {

    @Autowired
    private DocxTemplateService docxTemplateService;

    /**
     * Process a template file from classpath resources or uploaded file
     * 
     * @param templateName Name of template file in classpath resources/templates/
     * @param file Uploaded template file (alternative to templateName)
     * @param templateData JSON body containing variables map
     * @return Processed DOCX file as byte array
     */
    @PostMapping("/process")
    public ResponseEntity<byte[]> processTemplate(
            @RequestParam(value = "template", required = false) String templateName,
            @RequestParam(value = "file", required = false) MultipartFile file,
            @RequestBody TemplateData templateData) throws IOException {
        
        InputStream templateInputStream;
        
        if (file != null && !file.isEmpty()) {
            // Use uploaded file
            templateInputStream = file.getInputStream();
        } else if (templateName != null) {
            // Use template from classpath
            Resource resource = new ClassPathResource("templates/" + templateName);
            // Change localtion to a temporary directory instead of classpath
            
            if (!resource.exists()) {
                return ResponseEntity.badRequest().build();
            }
            templateInputStream = resource.getInputStream();
        } else {
            return ResponseEntity.badRequest().body("Either 'template' parameter or 'file' must be provided".getBytes());
        }

        byte[] processedDocument = docxTemplateService.processTemplatePreservingFormat(
                templateInputStream, 
                templateData.getVariables()
        );

        // Save to a file in the local filesystem
        // Generate UUID for the file name
        String fileName = UUID.randomUUID().toString() + ".docx";
        File outputFile = new File(fileName);
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            fos.write(processedDocument);
        }
        System.out.println("Processed document saved to: " + outputFile.getAbsolutePath());

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "processed-document.docx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(processedDocument);
    }

    /**
     * Health check endpoint
     */
    @GetMapping("/health")
    public ResponseEntity<String> health() {
        return ResponseEntity.ok("DOCX Template Processor Service is running");
    }
}

