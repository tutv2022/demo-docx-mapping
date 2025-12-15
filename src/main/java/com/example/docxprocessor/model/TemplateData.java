package com.example.docxprocessor.model;

import java.util.Map;

public class TemplateData {
    private Map<String, Object> variables;

    public TemplateData() {
    }

    public TemplateData(Map<String, Object> variables) {
        this.variables = variables;
    }

    public Map<String, Object> getVariables() {
        return variables;
    }

    public void setVariables(Map<String, Object> variables) {
        this.variables = variables;
    }
}

