package com.example.docxprocessor.model;

public class TemplateValidationIssue {
    /**
     * "ERROR" or "WARN"
     */
    private String severity;
    /**
     * Machine-friendly code (eg: JSON_PATH_NOT_FOUND, LOOP_UNMATCHED_END)
     */
    private String code;
    /**
     * Human-friendly message
     */
    private String message;
    /**
     * Optional context string (eg: the raw placeholder/tag)
     */
    private String context;

    public TemplateValidationIssue() {}

    public TemplateValidationIssue(String severity, String code, String message, String context) {
        this.severity = severity;
        this.code = code;
        this.message = message;
        this.context = context;
    }

    public String getSeverity() {
        return severity;
    }

    public void setSeverity(String severity) {
        this.severity = severity;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public String getContext() {
        return context;
    }

    public void setContext(String context) {
        this.context = context;
    }
}

