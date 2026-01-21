package com.example.docxprocessor.model;

public class TemplateValidationPlaceholder {
    /**
     * Raw tag in template (eg: {{$orderDate}}).
     */
    private String raw;
    /**
     * PLACEHOLDER | LOOP_ITEM | CHECKBOX | RADIO | UNKNOWN_TAG
     */
    private String type;
    /**
     * JSONPath expression used for validation (eg: $.orderDate).
     */
    private String jsonPath;
    /**
     * True if JSONPath syntax is valid/compilable.
     */
    private boolean jsonPathSyntaxValid;
    /**
     * True if JSONPath exists in sample JSON (read succeeds).
     */
    private boolean jsonPathFound;
    /**
     * True if placeholder is loop-dependent (contains [i]) and validation required substituting a sample index.
     */
    private boolean loopDependent;
    /**
     * Number of times this placeholder was seen.
     */
    private int occurrences;
    /**
     * Optional message.
     */
    private String message;

    public TemplateValidationPlaceholder() {}

    public TemplateValidationPlaceholder(String raw, String type) {
        this.raw = raw;
        this.type = type;
    }

    public String getRaw() {
        return raw;
    }

    public void setRaw(String raw) {
        this.raw = raw;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getJsonPath() {
        return jsonPath;
    }

    public void setJsonPath(String jsonPath) {
        this.jsonPath = jsonPath;
    }

    public boolean isJsonPathSyntaxValid() {
        return jsonPathSyntaxValid;
    }

    public void setJsonPathSyntaxValid(boolean jsonPathSyntaxValid) {
        this.jsonPathSyntaxValid = jsonPathSyntaxValid;
    }

    public boolean isJsonPathFound() {
        return jsonPathFound;
    }

    public void setJsonPathFound(boolean jsonPathFound) {
        this.jsonPathFound = jsonPathFound;
    }

    public boolean isLoopDependent() {
        return loopDependent;
    }

    public void setLoopDependent(boolean loopDependent) {
        this.loopDependent = loopDependent;
    }

    public int getOccurrences() {
        return occurrences;
    }

    public void setOccurrences(int occurrences) {
        this.occurrences = occurrences;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }
}

