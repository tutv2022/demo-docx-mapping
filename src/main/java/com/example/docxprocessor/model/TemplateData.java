package com.example.docxprocessor.model;

import java.util.Map;

public class TemplateData {
    private Map<String, String> variables;

    public TemplateData() {
    }

    public TemplateData(Map<String, String> variables) {
        this.variables = variables;
    }

    public Map<String, String> getVariables() {
        return variables;
    }

    public void setVariables(Map<String, String> variables) {
        this.variables = variables;
    }
}

