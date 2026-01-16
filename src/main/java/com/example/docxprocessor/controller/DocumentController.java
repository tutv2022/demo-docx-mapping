package com.example.docxprocessor.controller;

import com.example.docxprocessor.model.TemplateData;
import com.example.docxprocessor.service.DocxTemplateService;
import com.jayway.jsonpath.JsonPath;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

@RestController
@RequestMapping("/api/documents")
public class DocumentController {

    @Autowired
    private DocxTemplateService docxTemplateService;

    @Autowired
    private RestTemplate restTemplate;

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

    /**
     * Generate document from external API data
     * Calls external API, gets JSON response, and generates document using field paths from JSON
     * 
     * @param templateName Name of template file in classpath resources/templates/
     * @param file Uploaded template file (alternative to templateName)
     * @param apiUrl Optional API URL (defaults to http://localhost:8011/api/order)
     * @return Processed DOCX file as byte array
     */
    @PostMapping("/generate-from-api")
    public ResponseEntity<byte[]> generateFromApi(
            @RequestParam(value = "template", required = false) String templateName,
            @RequestParam(value = "file", required = false) MultipartFile file,
            @RequestParam(value = "apiUrl", required = false, defaultValue = "http://localhost:8011/api/order") String apiUrl) throws IOException {
        
        // Step 1: Call external API using RestTemplate and get response as JSON String
        String jsonResponse;
        try {
            jsonResponse = restTemplate.getForObject(apiUrl, String.class);
            if (jsonResponse == null || jsonResponse.isEmpty()) {
                return ResponseEntity.badRequest().body("API returned empty response".getBytes());
            }
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(("Failed to call external API: " + e.getMessage()).getBytes());
        }

        // Step 2: Parse JSON String using JSON-Path to extract field paths
        Map<String, Object> variables = new HashMap<>();
        try {
            // Parse JSON document using JSON-Path
            com.jayway.jsonpath.DocumentContext jsonContext = JsonPath.parse(jsonResponse);
            
            // Extract all fields from root level and nested levels
            // This will flatten the JSON structure for template variable access
            Object document = jsonContext.json();
            extractJsonPaths(document, "", variables);
            
            // Also add support for direct JSON-Path queries
            // Users can access values using JSON-Path expressions in templates
            // For example: {{$orderId}} or {{$customer.name}} or {{$items[0].name}}
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(("Failed to parse JSON response with JSON-Path: " + e.getMessage()).getBytes());
        }

        // Step 3: Get template input stream
        InputStream templateInputStream;
        
        if (file != null && !file.isEmpty()) {
            // Use uploaded file
            templateInputStream = file.getInputStream();
        } else if (templateName != null) {
            // Use template from classpath
            Resource resource = new ClassPathResource("templates/" + templateName);
            if (!resource.exists()) {
                return ResponseEntity.badRequest().body("Template file not found".getBytes());
            }
            templateInputStream = resource.getInputStream();
        } else {
            return ResponseEntity.badRequest().body("Either 'template' parameter or 'file' must be provided".getBytes());
        }

        // Step 4: Generate document using the parsed JSON data
        byte[] processedDocument = docxTemplateService.processTemplatePreservingFormat(
                templateInputStream, 
                variables
        );

        // Save to a file in the local filesystem
        String fileName = UUID.randomUUID().toString() + ".docx";
        File outputFile = new File(fileName);
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            fos.write(processedDocument);
        }
        System.out.println("Processed document saved to: " + outputFile.getAbsolutePath());

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "generated-document.docx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(processedDocument);
    }

    /**
     * Recursively extract JSON paths and values into a flat map
     * Supports nested objects and arrays
     * 
     * @param obj The JSON object/value to extract
     * @param prefix The current path prefix (e.g., "customer" or "items[0]")
     * @param variables The map to populate with extracted values
     */
    private void extractJsonPaths(Object obj, String prefix, Map<String, Object> variables) {
        if (obj == null) {
            if (!prefix.isEmpty()) {
                variables.put(prefix, null);
            }
            return;
        }

        if (obj instanceof Map) {
            @SuppressWarnings("unchecked")
            Map<String, Object> map = (Map<String, Object>) obj;
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                String key = entry.getKey();
                String newPrefix = prefix.isEmpty() ? key : prefix + "." + key;
                extractJsonPaths(entry.getValue(), newPrefix, variables);
            }
            // Also add the full path for nested objects
            if (!prefix.isEmpty()) {
                variables.put(prefix, obj);
            }
        } else if (obj instanceof java.util.List) {
            @SuppressWarnings("unchecked")
            java.util.List<Object> list = (java.util.List<Object>) obj;
            for (int i = 0; i < list.size(); i++) {
                String newPrefix = prefix.isEmpty() ? "[" + i + "]" : prefix + "[" + i + "]";
                extractJsonPaths(list.get(i), newPrefix, variables);
            }
            // Also add the full list
            if (!prefix.isEmpty()) {
                variables.put(prefix, obj);
            }
        } else {
            // Primitive value or string
            if (!prefix.isEmpty()) {
                variables.put(prefix, obj);
            }
        }
    }
}

