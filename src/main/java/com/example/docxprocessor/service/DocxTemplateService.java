package com.example.docxprocessor.service;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class DocxTemplateService {

    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)\\}");

    /**
     * Processes a DOCX template file by replacing placeholders with actual values
     * while preserving formatting
     * 
     * @param templateInputStream Input stream of the template DOCX file
     * @param data Map of placeholder names to replacement values
     * @return Byte array of the processed DOCX file
     * @throws IOException if file processing fails
     */
    public byte[] processTemplatePreservingFormat(InputStream templateInputStream, Map<String, String> data) throws IOException {
        try (XWPFDocument document = new XWPFDocument(templateInputStream)) {
            
            // Process paragraphs with better format preservation
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                replaceInParagraphPreservingFormat(paragraph, data);
            }

            // Process tables
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            replaceInParagraphPreservingFormat(paragraph, data);
                        }
                    }
                }
            }

            // Process headers
            for (XWPFHeader header : document.getHeaderList()) {
                for (XWPFParagraph paragraph : header.getParagraphs()) {
                    replaceInParagraphPreservingFormat(paragraph, data);
                }
            }

            // Process footers
            for (XWPFFooter footer : document.getFooterList()) {
                for (XWPFParagraph paragraph : footer.getParagraphs()) {
                    replaceInParagraphPreservingFormat(paragraph, data);
                }
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    /**
     * Enhanced replacement that better preserves formatting by working with individual runs
     */
    private void replaceInParagraphPreservingFormat(XWPFParagraph paragraph, Map<String, String> data) {
        if (paragraph.getRuns().isEmpty()) {
            return;
        }

        // Collect all text from runs
        StringBuilder fullText = new StringBuilder();
        for (XWPFRun run : paragraph.getRuns()) {
            String runText = run.getText(0);
            if (runText != null) {
                fullText.append(runText);
            }
        }

        String paragraphText = fullText.toString();
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(paragraphText);
        
        if (!matcher.find()) {
            return;
        }

        // Extract formatting properties BEFORE removing runs
        RunFormatting formatting = extractRunFormatting(paragraph.getRuns().get(0));

        // Build replacement text
        String replacedText = paragraphText;
        for (Map.Entry<String, String> entry : data.entrySet()) {
            String placeholder = "${" + entry.getKey() + "}";
            replacedText = replacedText.replace(placeholder, entry.getValue() != null ? entry.getValue() : "");
        }

        // Clear all runs
        int runsCount = paragraph.getRuns().size();
        for (int i = runsCount - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }

        // Create new run with replacement text and apply preserved formatting
        XWPFRun newRun = paragraph.createRun();
        newRun.setText(replacedText);
        applyRunFormatting(formatting, newRun);
    }

    /**
     * Data class to store run formatting properties
     */
    private static class RunFormatting {
        Boolean bold;
        Boolean italic;
        UnderlinePatterns underline;
        Boolean strike;
        String fontFamily;
        Integer fontSize;
        String color;
    }

    /**
     * Extracts formatting properties from a run BEFORE it's removed
     */
    @SuppressWarnings("deprecation")
    private RunFormatting extractRunFormatting(XWPFRun run) {
        RunFormatting formatting = new RunFormatting();
        if (run == null) {
            return formatting;
        }

        try {
            formatting.bold = run.isBold();
        } catch (Exception e) {
            formatting.bold = false;
        }

        try {
            formatting.italic = run.isItalic();
        } catch (Exception e) {
            formatting.italic = false;
        }

        try {
            formatting.underline = run.getUnderline();
        } catch (Exception e) {
            formatting.underline = UnderlinePatterns.NONE;
        }

        try {
            formatting.strike = run.isStrike();
        } catch (Exception e) {
            formatting.strike = false;
        }

        try {
            formatting.fontFamily = run.getFontFamily();
        } catch (Exception e) {
            formatting.fontFamily = null;
        }

        try {
            formatting.fontSize = run.getFontSize();
        } catch (Exception e) {
            formatting.fontSize = 0;
        }

        try {
            formatting.color = run.getColor();
        } catch (Exception e) {
            formatting.color = null;
        }

        return formatting;
    }

    /**
     * Applies formatting properties to a run
     */
    @SuppressWarnings("deprecation")
    private void applyRunFormatting(RunFormatting formatting, XWPFRun target) {
        if (formatting == null || target == null) {
            return;
        }

        if (formatting.bold != null) {
            target.setBold(formatting.bold);
        }

        if (formatting.italic != null) {
            target.setItalic(formatting.italic);
        }

        if (formatting.underline != null) {
            target.setUnderline(formatting.underline);
        }

        if (formatting.strike != null && formatting.strike) {
            target.setStrike(true);
        }

        if (formatting.fontFamily != null) {
            target.setFontFamily(formatting.fontFamily);
        }

        if (formatting.fontSize != null && formatting.fontSize > 0) {
            target.setFontSize(formatting.fontSize);
        }

        if (formatting.color != null) {
            target.setColor(formatting.color);
        }
    }
}

