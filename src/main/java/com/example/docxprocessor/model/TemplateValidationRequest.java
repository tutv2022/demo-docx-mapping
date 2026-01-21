package com.example.docxprocessor.model;

/**
 * Request body for template validation.
 *
 * Provide either:
 * - jsonSample: raw JSON string used to validate JSONPaths found in the template
 *
 * Or omit jsonSample and let the controller fetch JSON from an external API
 * (using apiUrl/method/token/requestBody query params).
 */
public class TemplateValidationRequest {
    private String jsonSample;

    public TemplateValidationRequest() {
    }

    public TemplateValidationRequest(String jsonSample) {
        this.jsonSample = jsonSample;
    }

    public String getJsonSample() {
        return jsonSample;
    }

    public void setJsonSample(String jsonSample) {
        this.jsonSample = jsonSample;
    }
}

