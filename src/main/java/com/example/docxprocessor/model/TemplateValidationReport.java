package com.example.docxprocessor.model;

import java.util.ArrayList;
import java.util.List;

public class TemplateValidationReport {
    private boolean valid;
    private int errorCount;
    private int warningCount;
    private List<TemplateValidationIssue> issues = new ArrayList<>();
    private List<TemplateValidationPlaceholder> placeholders = new ArrayList<>();
    private List<TemplateValidationLoop> loops = new ArrayList<>();

    public TemplateValidationReport() {}

    public boolean isValid() {
        return valid;
    }

    public void setValid(boolean valid) {
        this.valid = valid;
    }

    public int getErrorCount() {
        return errorCount;
    }

    public void setErrorCount(int errorCount) {
        this.errorCount = errorCount;
    }

    public int getWarningCount() {
        return warningCount;
    }

    public void setWarningCount(int warningCount) {
        this.warningCount = warningCount;
    }

    public List<TemplateValidationIssue> getIssues() {
        return issues;
    }

    public void setIssues(List<TemplateValidationIssue> issues) {
        this.issues = issues;
    }

    public List<TemplateValidationPlaceholder> getPlaceholders() {
        return placeholders;
    }

    public void setPlaceholders(List<TemplateValidationPlaceholder> placeholders) {
        this.placeholders = placeholders;
    }

    public List<TemplateValidationLoop> getLoops() {
        return loops;
    }

    public void setLoops(List<TemplateValidationLoop> loops) {
        this.loops = loops;
    }
}

