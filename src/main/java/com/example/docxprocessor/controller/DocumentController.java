package com.example.docxprocessor.controller;

import com.example.docxprocessor.model.TemplateData;
import com.example.docxprocessor.model.TemplateValidationReport;
import com.example.docxprocessor.model.TemplateValidationRequest;
import com.example.docxprocessor.service.Docx4jTemplateService;
import com.jayway.jsonpath.JsonPath;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;
import java.util.UUID;

@RestController
@RequestMapping("/api/documents")
public class DocumentController {

    @Autowired
    private Docx4jTemplateService docxTemplateService;

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
     * Supports GET, POST, and PUT methods with JWT token authentication
     * 
     * @param templateName Name of template file in classpath resources/templates/
     * @param file Uploaded template file (alternative to templateName)
     * @param apiUrl Optional API URL (defaults to http://localhost:8011/api/order)
     * @param method HTTP method (GET, POST, PUT) - defaults to GET
     * @param jwtToken JWT token for authentication (can be provided via Authorization header or token parameter)
     * @param requestBody Optional request body for POST/PUT requests
     * @param authorizationHeader Authorization header from request (alternative to jwtToken parameter)
     * @return Processed DOCX file as byte array
     */
    @PostMapping("/generate-from-api")
    public ResponseEntity<byte[]> generateFromApi(
            @RequestParam(value = "template", required = false) String templateName,
            @RequestParam(value = "file", required = false) MultipartFile file,
            @RequestParam(value = "apiUrl", required = false, defaultValue = "http://localhost:8011/api/order") String apiUrl,
            @RequestParam(value = "method", required = false, defaultValue = "GET") String method,
            @RequestParam(value = "token", required = false) String jwtToken,
            @RequestParam(value = "requestBody", required = false) String requestBodyJson,
            @RequestHeader(value = "Authorization", required = false) String authorizationHeader) throws IOException {
        
        // Extract JWT token from Authorization header or parameter
        String token = extractJwtToken(authorizationHeader, jwtToken);
        
        // Parse request body JSON if provided
        Map<String, Object> requestBody = null;
        if (requestBodyJson != null && !requestBodyJson.isEmpty()) {
            try {
                com.jayway.jsonpath.DocumentContext bodyContext = JsonPath.parse(requestBodyJson);
                Object bodyObj = bodyContext.json();
                if (bodyObj instanceof Map) {
                    @SuppressWarnings("unchecked")
                    Map<String, Object> bodyMap = (Map<String, Object>) bodyObj;
                    requestBody = bodyMap;
                }
            } catch (Exception e) {
                return ResponseEntity.badRequest().body(("Invalid request body JSON: " + e.getMessage()).getBytes());
            }
        }
        
        // Step 1: Call external API using RestTemplate with specified method and JWT token
        String jsonResponse;
        try {
            jsonResponse = callExternalApi(apiUrl, method, token, requestBody);
            if (jsonResponse == null || jsonResponse.isEmpty()) {
                return ResponseEntity.badRequest().body("API returned empty response".getBytes());
            }
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(("Failed to call external API: " + e.getMessage()).getBytes());
        }

        // Step 2: Get template input stream
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

        // Step 3: Generate document using JSON-Path to extract values directly from JSON
        // Placeholders like {{$orderId}} or {{$customer.name}} will use JSON-Path expressions
        byte[] processedDocument;
        try {
            processedDocument = docxTemplateService.processTemplateWithJsonPath(
                    templateInputStream, 
                    jsonResponse
            );
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(("Failed to process template with JSON-Path: " + e.getMessage()).getBytes());
        }

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
     * Generate PDF document from external API data.
     * Same as /generate-from-api but exports PDF.
     */
    @PostMapping("/generate-pdf-from-api")
    public ResponseEntity<byte[]> generatePdfFromApi(
            @RequestParam(value = "template", required = false) String templateName,
            @RequestParam(value = "file", required = false) MultipartFile file,
            @RequestParam(value = "apiUrl", required = false, defaultValue = "http://localhost:8011/api/order") String apiUrl,
            @RequestParam(value = "method", required = false, defaultValue = "GET") String method,
            @RequestParam(value = "token", required = false) String jwtToken,
            @RequestParam(value = "requestBody", required = false) String requestBodyJson,
            @RequestHeader(value = "Authorization", required = false) String authorizationHeader) throws IOException {

        String token = extractJwtToken(authorizationHeader, jwtToken);

        Map<String, Object> requestBody = null;
        if (requestBodyJson != null && !requestBodyJson.isEmpty()) {
            try {
                com.jayway.jsonpath.DocumentContext bodyContext = JsonPath.parse(requestBodyJson);
                Object bodyObj = bodyContext.json();
                if (bodyObj instanceof Map) {
                    @SuppressWarnings("unchecked")
                    Map<String, Object> bodyMap = (Map<String, Object>) bodyObj;
                    requestBody = bodyMap;
                }
            } catch (Exception e) {
                return ResponseEntity.badRequest().body(("Invalid request body JSON: " + e.getMessage()).getBytes());
            }
        }

        String jsonResponse;
        try {
            jsonResponse = callExternalApi(apiUrl, method, token, requestBody);
            if (jsonResponse == null || jsonResponse.isEmpty()) {
                return ResponseEntity.badRequest().body("API returned empty response".getBytes());
            }
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(("Failed to call external API: " + e.getMessage()).getBytes());
        }

        InputStream templateInputStream;
        if (file != null && !file.isEmpty()) {
            templateInputStream = file.getInputStream();
        } else if (templateName != null) {
            Resource resource = new ClassPathResource("templates/" + templateName);
            if (!resource.exists()) {
                return ResponseEntity.badRequest().body("Template file not found".getBytes());
            }
            templateInputStream = resource.getInputStream();
        } else {
            return ResponseEntity.badRequest().body("Either 'template' parameter or 'file' must be provided".getBytes());
        }

        byte[] pdfBytes;
        try {
            pdfBytes = docxTemplateService.processTemplateWithJsonPathToPdf(templateInputStream, jsonResponse);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(("Failed to generate PDF: " + e.getMessage()).getBytes());
        }

        // Save to a file in the local filesystem (same as DOCX flow)
        String fileName = UUID.randomUUID().toString() + ".pdf";
        File outputFile = new File(fileName);
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            fos.write(pdfBytes);
        }
        System.out.println("Generated PDF saved to: " + outputFile.getAbsolutePath());

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDisposition(ContentDisposition.attachment().filename("generated-document.pdf").build());

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }

    /**
     * Validate a template against supported syntax and sample JSON (JSONPath).
     *
     * Provide the template via:
     * - template: classpath resource under templates/
     * - file: multipart upload
     *
     * Provide JSON via:
     * - request body { "jsonSample": "<raw json>" }
     *   OR
     * - external API params (apiUrl/method/token/requestBody), same as /generate-from-api
     */
    @PostMapping("/validate-template")
    public ResponseEntity<?> validateTemplate(
            @RequestParam(value = "template", required = false) String templateName,
            @RequestParam(value = "file", required = false) MultipartFile file,
            @RequestParam(value = "apiUrl", required = false, defaultValue = "http://localhost:8011/api/order") String apiUrl,
            @RequestParam(value = "method", required = false, defaultValue = "GET") String method,
            @RequestParam(value = "token", required = false) String jwtToken,
            @RequestParam(value = "requestBody", required = false) String requestBodyJson,
            @RequestHeader(value = "Authorization", required = false) String authorizationHeader,
            @RequestBody(required = false) TemplateValidationRequest validationRequest) throws IOException {

        // JSON source
        String jsonResponse = null;
        if (validationRequest != null && validationRequest.getJsonSample() != null && !validationRequest.getJsonSample().isEmpty()) {
            jsonResponse = validationRequest.getJsonSample();
        } else {
            String token = extractJwtToken(authorizationHeader, jwtToken);

            Map<String, Object> requestBody = null;
            if (requestBodyJson != null && !requestBodyJson.isEmpty()) {
                try {
                    com.jayway.jsonpath.DocumentContext bodyContext = JsonPath.parse(requestBodyJson);
                    Object bodyObj = bodyContext.json();
                    if (bodyObj instanceof Map) {
                        @SuppressWarnings("unchecked")
                        Map<String, Object> bodyMap = (Map<String, Object>) bodyObj;
                        requestBody = bodyMap;
                    }
                } catch (Exception e) {
                    return ResponseEntity.badRequest().body("Invalid request body JSON: " + e.getMessage());
                }
            }

            try {
                jsonResponse = callExternalApi(apiUrl, method, token, requestBody);
                if (jsonResponse == null || jsonResponse.isEmpty()) {
                    return ResponseEntity.badRequest().body("API returned empty response");
                }
            } catch (Exception e) {
                return ResponseEntity.badRequest().body("Failed to call external API: " + e.getMessage());
            }
        }

        // Template source
        InputStream templateInputStream;
        if (file != null && !file.isEmpty()) {
            templateInputStream = file.getInputStream();
        } else if (templateName != null) {
            Resource resource = new ClassPathResource("templates/" + templateName);
            if (!resource.exists()) {
                return ResponseEntity.badRequest().body("Template file not found");
            }
            templateInputStream = resource.getInputStream();
        } else {
            return ResponseEntity.badRequest().body("Either 'template' parameter or 'file' must be provided");
        }

        try {
            TemplateValidationReport report = docxTemplateService.validateTemplateWithJsonPath(templateInputStream, jsonResponse);
            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_JSON)
                    .body(report);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body("Failed to validate template: " + e.getMessage());
        }
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

    /**
     * Extract JWT token from Authorization header or token parameter
     * Supports "Bearer <token>" format in Authorization header
     * 
     * @param authorizationHeader Authorization header value
     * @param jwtToken JWT token from parameter
     * @return JWT token string
     */
    private String extractJwtToken(String authorizationHeader, String jwtToken) {
        if (jwtToken != null && !jwtToken.isEmpty()) {
            return jwtToken;
        }
        
        if (authorizationHeader != null && !authorizationHeader.isEmpty()) {
            // Remove "Bearer " prefix if present
            if (authorizationHeader.startsWith("Bearer ")) {
                return authorizationHeader.substring(7);
            }
            return authorizationHeader;
        }
        
        return null;
    }

    /**
     * Call external API with specified HTTP method and JWT token
     * 
     * @param apiUrl The API URL to call
     * @param method HTTP method (GET, POST, PUT)
     * @param jwtToken JWT token for authentication
     * @param requestBody Request body for POST/PUT requests
     * @return JSON response as String
     * @throws IllegalArgumentException if HTTP method is not supported
     */
    private String callExternalApi(String apiUrl, String method, String jwtToken, Map<String, Object> requestBody) {
        // Validate HTTP method
        String upperMethod = method.toUpperCase();
        if (!upperMethod.equals("GET") && !upperMethod.equals("POST") && !upperMethod.equals("PUT")) {
            throw new IllegalArgumentException("Unsupported HTTP method: " + method + ". Supported methods: GET, POST, PUT");
        }
        
        HttpMethod httpMethod = HttpMethod.valueOf(upperMethod);
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON);
        
        // Add JWT token to Authorization header if provided
        if (jwtToken != null && !jwtToken.isEmpty()) {
            headers.setBearerAuth(jwtToken);
        }
        
        HttpEntity<?> entity;
        
        // Create request entity with body for POST/PUT
        if ((httpMethod == HttpMethod.POST || httpMethod == HttpMethod.PUT) && requestBody != null) {
            entity = new HttpEntity<>(requestBody, headers);
        } else {
            entity = new HttpEntity<>(headers);
        }
        
        ResponseEntity<String> response = restTemplate.exchange(
                apiUrl,
                httpMethod,
                entity,
                String.class
        );
        
        return response.getBody();
    }
}

