package com.example.docxprocessor.model;

public class TemplateValidationLoop {
    /**
     * TABLE or PARAGRAPH
     */
    private String scope;
    /**
     * Array name referenced by loop start marker (eg: orderItems)
     */
    private String arrayName;
    /**
     * True if the loop start marker was paired with an end marker.
     */
    private boolean paired;
    /**
     * Array length from sample JSON if available (null when not readable).
     */
    private Integer arrayLength;
    /**
     * Optional message.
     */
    private String message;

    public TemplateValidationLoop() {}

    public TemplateValidationLoop(String scope, String arrayName) {
        this.scope = scope;
        this.arrayName = arrayName;
    }

    public String getScope() {
        return scope;
    }

    public void setScope(String scope) {
        this.scope = scope;
    }

    public String getArrayName() {
        return arrayName;
    }

    public void setArrayName(String arrayName) {
        this.arrayName = arrayName;
    }

    public boolean isPaired() {
        return paired;
    }

    public void setPaired(boolean paired) {
        this.paired = paired;
    }

    public Integer getArrayLength() {
        return arrayLength;
    }

    public void setArrayLength(Integer arrayLength) {
        this.arrayLength = arrayLength;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }
}

