package com.example.docxprocessor.service;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class DocxTemplateService {

    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)\\}");
    private static final Pattern CHECKBOX_PATTERN = Pattern.compile("\\$\\{checkbox:([^:]+):([^}]+)\\}");
    private static final Pattern RADIO_PATTERN = Pattern.compile("\\$\\{radio:([^:]+):([^}]+)\\}");
    // Loop patterns - escape # and . properly
    private static final Pattern LOOP_START_PATTERN = Pattern.compile("\\$\\{#loop\\.([^}]+)\\}");
    private static final Pattern LOOP_END_PATTERN = Pattern.compile("\\$\\{#loop\\}");

    /**
     * Processes a DOCX template file by replacing placeholders with actual values
     * while preserving formatting
     * 
     * @param templateInputStream Input stream of the template DOCX file
     * @param data Map of placeholder names to replacement values
     * @return Byte array of the processed DOCX file
     * @throws IOException if file processing fails
     */
    public byte[] processTemplatePreservingFormat(InputStream templateInputStream, Map<String, Object> data) throws IOException {
        try (XWPFDocument document = new XWPFDocument(templateInputStream)) {
            
            // Step 1: Process loops first (they can contain placeholders)
            processLoops(document, data);
            
            // Step 2: Process paragraphs - unified processing for checkboxes and regular placeholders
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                processAllPlaceholders(paragraph, data);
            }

            // Process tables
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            processAllPlaceholders(paragraph, data);
                        }
                    }
                }
            }

            // Process headers
            for (XWPFHeader header : document.getHeaderList()) {
                for (XWPFParagraph paragraph : header.getParagraphs()) {
                    processAllPlaceholders(paragraph, data);
                }
            }

            // Process footers
            for (XWPFFooter footer : document.getFooterList()) {
                for (XWPFParagraph paragraph : footer.getParagraphs()) {
                    processAllPlaceholders(paragraph, data);
                }
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    /**
     * Processes loop blocks in the document
     * Syntax: ${#loop.variableName} ... content ... ${#loop}
     */
    private void processLoops(XWPFDocument document, Map<String, Object> data) {
        List<XWPFParagraph> paragraphs = new ArrayList<>(document.getParagraphs());
        
        if (paragraphs.isEmpty()) {
            return;
        }
        
        // Keep processing until no more loops are found
        int maxIterations = 100; // Safety limit
        int iteration = 0;
        
        while (iteration < maxIterations) {
            iteration++;
            boolean foundLoop = false;
            
            // Scan all paragraphs from end to beginning to find loop end markers
            for (int i = paragraphs.size() - 1; i >= 0; i--) {
                XWPFParagraph paragraph = paragraphs.get(i);
                if (paragraph == null) {
                    continue;
                }
                
                String text = getParagraphText(paragraph);
                
                // Debug: Print paragraph text to see what we're checking
                // System.out.println("Checking paragraph " + i + ": " + text);
                
                // Check if this paragraph contains a loop end marker
                Matcher endMatcher = LOOP_END_PATTERN.matcher(text);
                if (endMatcher.find()) {
                    // Found end marker, find corresponding start marker
                    int startIndex = findLoopStartMarker(paragraphs, i);
                    if (startIndex != -1 && startIndex < i) {
                        // Process this loop block
                        processLoopBlock(document, paragraphs, startIndex, i, data);
                        foundLoop = true;
                        // Refresh paragraph list after modification
                        paragraphs = new ArrayList<>(document.getParagraphs());
                        break; // Restart scanning
                    }
                }
            }
            
            if (!foundLoop) {
                break; // No more loops found
            }
        }
    }

    /**
     * Gets full text from a paragraph by collecting all runs
     * This ensures we capture placeholders that might be split across multiple runs
     */
    private String getParagraphText(XWPFParagraph paragraph) {
        if (paragraph == null) {
            return "";
        }
        
        // Try to get text from all runs first
        if (!paragraph.getRuns().isEmpty()) {
            StringBuilder fullText = new StringBuilder();
            for (XWPFRun run : paragraph.getRuns()) {
                if (run == null) {
                    continue;
                }
                
                // Safely get text from run - handle IndexOutOfBoundsException
                try {
                    String runText = run.getText(0);
                    if (runText != null) {
                        fullText.append(runText);
                    }
                } catch (IndexOutOfBoundsException e) {
                    // Run has no text at position 0, skip it
                    // This is normal for runs that only contain formatting or other elements
                    continue;
                } catch (Exception e) {
                    // Any other exception, skip this run
                    continue;
                }
            }
            String result = fullText.toString();
            if (!result.isEmpty()) {
                return result;
            }
        }
        
        // Fallback to paragraph.getText() - this is safer and handles empty runs
        try {
            String text = paragraph.getText();
            return text != null ? text : "";
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * Checks if text contains a loop end marker
     */
    private boolean containsLoopEnd(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        return LOOP_END_PATTERN.matcher(text).find();
    }

    /**
     * Checks if text contains a loop start marker
     */
    private boolean containsLoopStart(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        return LOOP_START_PATTERN.matcher(text).find();
    }

    /**
     * Finds the start marker for a loop ending at endIndex
     * Searches backwards from endIndex to find the matching start marker
     */
    private int findLoopStartMarker(List<XWPFParagraph> paragraphs, int endIndex) {
        if (endIndex < 0 || endIndex >= paragraphs.size()) {
            return -1;
        }
        
        int depth = 1; // Start with depth 1 (we found one end marker)
        
        // Search backwards from endIndex
        for (int i = endIndex - 1; i >= 0; i--) {
            XWPFParagraph paragraph = paragraphs.get(i);
            if (paragraph == null) {
                continue;
            }
            
            String text = getParagraphText(paragraph);
            if (text == null || text.isEmpty()) {
                continue;
            }
            
            // Check for end markers (nested loops) - must reset matcher each time
            Matcher endMatcher = LOOP_END_PATTERN.matcher(text);
            if (endMatcher.find()) {
                depth++;
            }
            
            // Check for start markers - must reset matcher each time
            Matcher startMatcher = LOOP_START_PATTERN.matcher(text);
            if (startMatcher.find()) {
                depth--;
                if (depth == 0) {
                    // Found the matching start marker
                    return i;
                }
            }
        }
        
        return -1; // Start marker not found
    }

    /**
     * Processes a single loop block
     */
    @SuppressWarnings("unchecked")
    private void processLoopBlock(XWPFDocument document, List<XWPFParagraph> paragraphs, 
                                  int startIndex, int endIndex, Map<String, Object> data) {
        XWPFParagraph startPara = paragraphs.get(startIndex);
        // Use getParagraphText to get full text from all runs
        String startText = getParagraphText(startPara);
        
        Matcher startMatcher = LOOP_START_PATTERN.matcher(startText);
        if (!startMatcher.find()) {
            return;
        }
        
        String loopVariable = startMatcher.group(1); // e.g., "orderItems"
        
        // Get the list from data
        Object listObj = data.get(loopVariable);
        if (!(listObj instanceof List)) {
            // Not a list, remove the loop markers
            removeLoopMarkers(paragraphs, startIndex, endIndex);
            return;
        }
        
        List<Map<String, Object>> items = (List<Map<String, Object>>) listObj;
        if (items.isEmpty()) {
            // Empty list, remove the loop block
            removeLoopBlock(paragraphs, startIndex, endIndex);
            return;
        }
        
        // Extract content between markers (excluding the markers themselves)
        List<XWPFParagraph> loopContent = new ArrayList<>();
        for (int i = startIndex + 1; i < endIndex; i++) {
            loopContent.add(paragraphs.get(i));
        }
        
        // IMPORTANT: Clone paragraph XML BEFORE removing them from document
        // This prevents XmlValueDisconnectedException
        List<String> loopContentXml = new ArrayList<>();
        for (XWPFParagraph para : loopContent) {
            // Get XML text while paragraph is still connected
            String paraXml = para.getCTP().xmlText();
            loopContentXml.add(paraXml);
        }
        
        // Remove the loop markers from start and end paragraphs
        removePlaceholderFromParagraph(startPara, LOOP_START_PATTERN);
        XWPFParagraph endPara = paragraphs.get(endIndex);
        removePlaceholderFromParagraph(endPara, LOOP_END_PATTERN);
        
        // Get the CTBody to manipulate XML directly
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody body = document.getDocument().getBody();
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP startParaCT = startPara.getCTP();
        
        // Find the index of startPara in the body's P array by comparing XML
        int startParaIndex = -1;
        String startParaXml = startParaCT.xmlText();
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i).xmlText().equals(startParaXml)) {
                startParaIndex = i;
                break;
            }
        }
        
        if (startParaIndex == -1) {
            return; // Couldn't find the paragraph
        }
        
        // Remove the original loop content paragraphs and end marker BEFORE inserting new ones
        // This avoids index shifting issues
        // Remove from end to start to maintain correct indices
        int endParaIndex = -1;
        String endParaXml = endPara.getCTP().xmlText();
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i).xmlText().equals(endParaXml)) {
                endParaIndex = i;
                break;
            }
        }
        
        if (endParaIndex == -1 || endParaIndex <= startParaIndex) {
            // End marker not found or invalid - can't proceed safely
            return;
        }
        
        // Remove end marker paragraph first
        body.removeP(endParaIndex);
        // After removal, endParaIndex is now one less, but we don't need it anymore
        
        // Remove loop content paragraphs (between start and end, excluding start itself)
        // Remove from end to start to avoid index shifting
        // After removing endPara, the last content paragraph is at endParaIndex - 1
        for (int i = endParaIndex - 1; i > startParaIndex; i--) {
            body.removeP(i);
        }
        
        // Create and insert paragraphs for each item
        int currentInsertPos = startParaIndex + 1;
        
        for (Map<String, Object> item : items) {
            // Process each paragraph XML in the loop content
            for (String paraXml : loopContentXml) {
                // Parse the XML to create a new CT element
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP clonedCT;
                try {
                    clonedCT = org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP.Factory.parse(paraXml);
                } catch (org.apache.xmlbeans.XmlException e) {
                    // If parsing fails, skip this paragraph
                    continue;
                }
                
                // Insert new paragraph at position and replace it with our cloned content
                body.insertNewP(currentInsertPos);
                // Replace the empty paragraph with our cloned one
                body.setPArray(currentInsertPos, clonedCT);
                
                // Get the paragraph we just inserted
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP newParaCT = body.getPArray(currentInsertPos);
                
                // Create XWPFParagraph wrapper to process placeholders
                XWPFParagraph newPara = new XWPFParagraph(newParaCT, document);
                
                // Process placeholders with item data
                // This must work correctly - if paragraph has no runs, it means it's empty and we skip it
                if (!newPara.getRuns().isEmpty()) {
                    processAllPlaceholders(newPara, item);
                } else {
                    // If no runs, check if there's text in the CT element that needs processing
                    String paraText = getParagraphText(newPara);
                    if (paraText != null && !paraText.trim().isEmpty()) {
                        // There's text but no runs - this shouldn't happen with a cloned paragraph
                        // But if it does, create a run and process it
                        XWPFRun run = newPara.createRun();
                        run.setText(paraText);
                        processAllPlaceholders(newPara, item);
                    }
                }
                
                // Move insertion position forward
                currentInsertPos++;
            }
        }
    }


    /**
     * Finds the insertion position for a paragraph in the document body
     */
    private int findInsertPosition(XWPFDocument document, XWPFParagraph targetPara) {
        List<IBodyElement> bodyElements = document.getBodyElements();
        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement element = bodyElements.get(i);
            if (element == targetPara) {
                return i;
            }
        }
        return -1;
    }


    /**
     * Removes the entire loop block (for empty lists)
     */
    private void removeLoopBlock(List<XWPFParagraph> paragraphs, int startIndex, int endIndex) {
        // This would need to remove from document, which is complex
        // For now, just remove markers
        removeLoopMarkers(paragraphs, startIndex, endIndex);
    }

    /**
     * Removes loop markers from paragraphs
     */
    private void removeLoopMarkers(List<XWPFParagraph> paragraphs, int startIndex, int endIndex) {
        removePlaceholderFromParagraph(paragraphs.get(startIndex), LOOP_START_PATTERN);
        removePlaceholderFromParagraph(paragraphs.get(endIndex), LOOP_END_PATTERN);
    }

    /**
     * Removes a placeholder pattern from a paragraph
     */
    private void removePlaceholderFromParagraph(XWPFParagraph paragraph, Pattern pattern) {
        String text = paragraph.getText();
        Matcher matcher = pattern.matcher(text);
        if (matcher.find()) {
            String newText = text.replace(matcher.group(0), "").trim();
            // Clear and recreate runs
            int runsCount = paragraph.getRuns().size();
            for (int i = runsCount - 1; i >= 0; i--) {
                paragraph.removeRun(i);
            }
            if (!newText.isEmpty()) {
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(newText);
            }
        }
    }

    /**
     * Clones a paragraph with all its content and formatting
     */
    private XWPFParagraph cloneParagraph(XWPFDocument document, XWPFParagraph source) {
        XWPFParagraph target = document.createParagraph();
        
        // Copy paragraph properties
        if (source.getAlignment() != null) {
            target.setAlignment(source.getAlignment());
        }
        target.setIndentationLeft(source.getIndentationLeft());
        target.setIndentationRight(source.getIndentationRight());
        target.setSpacingBefore(source.getSpacingBefore());
        target.setSpacingAfter(source.getSpacingAfter());
        
        // Copy runs
        for (XWPFRun sourceRun : source.getRuns()) {
            XWPFRun targetRun = target.createRun();
            copyRun(sourceRun, targetRun);
        }
        
        return target;
    }

    /**
     * Copies run content and formatting
     */
    @SuppressWarnings("deprecation")
    private void copyRun(XWPFRun source, XWPFRun target) {
        String text = source.getText(0);
        if (text != null) {
            target.setText(text);
        }
        
        target.setBold(source.isBold());
        target.setItalic(source.isItalic());
        target.setUnderline(source.getUnderline());
        
        if (source.getFontFamily() != null) {
            target.setFontFamily(source.getFontFamily());
        }
        
        if (source.getFontSize() > 0) {
            target.setFontSize(source.getFontSize());
        }
        
        if (source.getColor() != null) {
            target.setColor(source.getColor());
        }
        
        // Copy highlight if present
        try {
            if (source.getCTR() != null && source.getCTR().getRPr() != null) {
                String rprXml = source.getCTR().getRPr().xmlText();
                String highlightValue = extractHighlightFromXml(rprXml);
                if (highlightValue != null) {
                    if (target.getCTR().getRPr() == null) {
                        target.getCTR().addNewRPr();
                    }
                    // Copy highlight using XML cursor
                    org.apache.xmlbeans.XmlCursor cursor = target.getCTR().getRPr().newCursor();
                    cursor.toEndToken();
                    cursor.insertElement(
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                        "highlight"
                    );
                    cursor.insertAttributeWithValue(
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                        "val",
                        highlightValue
                    );
                    cursor.dispose();
                }
            }
        } catch (Exception e) {
            // Ignore if highlight cannot be copied
        }
    }

    /**
     * Processes radio buttons, checkboxes, and regular placeholders in a single pass
     * This prevents duplication and ensures proper replacement
     */
    private void processAllPlaceholders(XWPFParagraph paragraph, Map<String, Object> data) {
        if (paragraph.getRuns().isEmpty()) {
            return;
        }

        // Collect all text from runs using safe method
        String paragraphText = getParagraphText(paragraph);
        
        // Check if paragraph has any placeholders
        Matcher radioMatcher = RADIO_PATTERN.matcher(paragraphText);
        Matcher checkboxMatcher = CHECKBOX_PATTERN.matcher(paragraphText);
        Matcher regularMatcher = PLACEHOLDER_PATTERN.matcher(paragraphText);
        
        boolean hasRadios = radioMatcher.find();
        boolean hasCheckboxes = checkboxMatcher.find();
        boolean hasRegularPlaceholders = regularMatcher.find();
        
        if (!hasRadios && !hasCheckboxes && !hasRegularPlaceholders) {
            return; // No placeholders in this paragraph
        }

        String replacedText = paragraphText;

        // Step 1: Process radio button placeholders first
        if (hasRadios) {
            radioMatcher.reset();
            Map<String, String> radioReplacements = new HashMap<>();
            
            while (radioMatcher.find()) {
                String radioGroup = radioMatcher.group(1); // e.g., "gender"
                String radioValue = radioMatcher.group(2); // e.g., "male" or "female"
                String fullPlaceholder = radioMatcher.group(0); // e.g., "${radio:gender:male}"
                
                // Get the actual value for this group from data
                Object groupValueObj = data.get(radioGroup);
                String groupValue = groupValueObj != null ? groupValueObj.toString() : null;
                
                // Determine if this radio button should be selected
                boolean isSelected = radioValue.equalsIgnoreCase(groupValue);
                
                // Replace with radio button symbol (● = checked, ○ = unchecked)
                //String radioSymbol = isSelected ? "● " : "○ ";
                String radioSymbol = isSelected ? "◉" : "○";
                radioReplacements.put(fullPlaceholder, radioSymbol);
            }
            
            // Apply radio button replacements
            for (Map.Entry<String, String> entry : radioReplacements.entrySet()) {
                replacedText = replacedText.replace(entry.getKey(), entry.getValue());
            }
        }

        // Step 2: Process checkbox placeholders
        if (hasCheckboxes) {
            checkboxMatcher.reset();
            Map<String, String> checkboxReplacements = new HashMap<>();
            
            while (checkboxMatcher.find()) {
                String checkboxGroup = checkboxMatcher.group(1); // e.g., "gender"
                String checkboxValue = checkboxMatcher.group(2); // e.g., "male" or "female"
                String fullPlaceholder = checkboxMatcher.group(0); // e.g., "${checkbox:gender:male}"
                
                // Get the actual value for this group from data
                Object groupValueObj = data.get(checkboxGroup);
                String groupValue = groupValueObj != null ? groupValueObj.toString() : null;
                
                // Determine if this checkbox should be checked
                boolean isChecked = checkboxValue.equalsIgnoreCase(groupValue);
                
                // Replace with checkbox symbol
                String checkboxSymbol = isChecked ? "☒" : "☐";
                checkboxReplacements.put(fullPlaceholder, checkboxSymbol);
            }
            
            // Apply checkbox replacements
            for (Map.Entry<String, String> entry : checkboxReplacements.entrySet()) {
                replacedText = replacedText.replace(entry.getKey(), entry.getValue());
            }
        }

        // Step 3: Process regular placeholders
        if (hasRegularPlaceholders) {
            for (Map.Entry<String, Object> entry : data.entrySet()) {
                String placeholder = "${" + entry.getKey() + "}";
                String value = entry.getValue() != null ? entry.getValue().toString() : "";
                replacedText = replacedText.replace(placeholder, value);
            }
        }

        // Only update if text changed
        if (!replacedText.equals(paragraphText)) {
            // Extract formatting before removing runs
            RunFormatting formatting = extractRunFormatting(paragraph.getRuns().get(0));
            
            // Clear all runs
            int runsCount = paragraph.getRuns().size();
            for (int i = runsCount - 1; i >= 0; i--) {
                paragraph.removeRun(i);
            }
            
            // Create new run with replaced text
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(replacedText);
            applyRunFormatting(formatting, newRun);
        }
    }

    /**
     * Processes checkbox placeholders before regular placeholders
     * Handles patterns like: ${checkbox:gender:male} or ${checkbox:gender:female}
     * Automatically sets ☑ for checked and ☐ for unchecked based on group value
     */
    private void processCheckboxPlaceholders(XWPFParagraph paragraph, Map<String, String> data) {
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
        Matcher checkboxMatcher = CHECKBOX_PATTERN.matcher(paragraphText);
        
        if (!checkboxMatcher.find()) {
            return; // No checkboxes in this paragraph
        }

        // Build replacement map for checkboxes
        Map<String, String> checkboxReplacements = new HashMap<>();
        checkboxMatcher.reset(); // Reset matcher to start from beginning
        
        while (checkboxMatcher.find()) {
            String checkboxGroup = checkboxMatcher.group(1); // e.g., "gender"
            String checkboxValue = checkboxMatcher.group(2); // e.g., "male" or "female"
            String fullPlaceholder = checkboxMatcher.group(0); // e.g., "${checkbox:gender:male}"
            
            // Get the actual value for this group from data
            String groupValue = data.get(checkboxGroup); // e.g., data.get("gender") -> "male"
            
            // Determine if this checkbox should be checked
            boolean isChecked = checkboxValue.equalsIgnoreCase(groupValue);
            
            // Replace with checkbox symbol followed by a space for better formatting
//            String checkboxSymbol = isChecked ? "☑ " : "☐ ";
//            String checkboxSymbol = isChecked ? "V" : "x";
            // Wingdings: ☐ = 0x6F, ☑ = 0xFE
            String checkboxSymbol = isChecked ? "\u00FE " : "\u006F ";
            checkboxReplacements.put(fullPlaceholder, checkboxSymbol);
        }

        // If we have replacements, update the paragraph
        if (!checkboxReplacements.isEmpty()) {
            // Apply checkbox replacements
            String replacedText = paragraphText;
            for (Map.Entry<String, String> entry : checkboxReplacements.entrySet()) {
                replacedText = replacedText.replace(entry.getKey(), entry.getValue());
            }

            // Extract formatting before removing runs
            RunFormatting formatting = extractRunFormatting(paragraph.getRuns().get(0));
            
            // Clear all runs
            int runsCount = paragraph.getRuns().size();
            for (int i = runsCount - 1; i >= 0; i--) {
                paragraph.removeRun(i);
            }
            
            // Create new run with replaced text
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(replacedText);
            applyRunFormatting(formatting, newRun);
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
        String highlight;  // Background color/highlighting
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

        // Extract highlight/background color by copying the entire RPr XML
        try {
            if (run.getCTR() != null && run.getCTR().getRPr() != null) {
                // Store the entire RPr XML structure as a string
                formatting.highlight = run.getCTR().getRPr().xmlText();
            }
        } catch (Exception e) {
            formatting.highlight = null;
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

        // Apply highlight/background color by copying from source RPr XML
        if (formatting.highlight != null && !formatting.highlight.isEmpty()) {
            try {
                // Extract highlight value from the stored RPr XML
                String highlightValue = extractHighlightFromXml(formatting.highlight);
                if (highlightValue != null) {
                    // Ensure RPr exists on target
                    if (target.getCTR().getRPr() == null) {
                        target.getCTR().addNewRPr();
                    }
                    
                    // Use XML cursor to add highlight element
                    org.apache.xmlbeans.XmlCursor cursor = target.getCTR().getRPr().newCursor();
                    // Move to end of RPr to insert highlight
                    cursor.toEndToken();
                    // Insert highlight element
                    cursor.insertElement(
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                        "highlight"
                    );
                    // Set the val attribute
                    cursor.insertAttributeWithValue(
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                        "val",
                        highlightValue
                    );
                    cursor.dispose();
                }
            } catch (Exception e) {
                // Try alternative approach using DOM directly
                try {
                    String highlightValue = extractHighlightFromXml(formatting.highlight);
                    if (highlightValue != null && target.getCTR().getRPr() != null) {
                        org.w3c.dom.Document doc = target.getCTR().getRPr().getDomNode().getOwnerDocument();
                        org.w3c.dom.Element highlightElement = doc.createElementNS(
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                            "w:highlight"
                        );
                        highlightElement.setAttributeNS(
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                            "w:val",
                            highlightValue
                        );
                        target.getCTR().getRPr().getDomNode().appendChild(highlightElement);
                    }
                } catch (Exception ex) {
                    // Ignore if highlight cannot be applied - document will still work
                }
            }
        }
    }
    
    /**
     * Helper method to extract highlight value from RPr XML string
     */
    private String extractHighlightFromXml(String rprXml) {
        if (rprXml == null || rprXml.isEmpty()) {
            return null;
        }
        try {
            // Look for highlight element - try with namespace prefix first
            int highlightElementStart = rprXml.indexOf("<w:highlight");
            if (highlightElementStart < 0) {
                // Try without namespace prefix
                highlightElementStart = rprXml.indexOf("<highlight");
            }
            
            if (highlightElementStart >= 0) {
                // Find the val attribute - try with namespace prefix first
                int valStart = rprXml.indexOf("w:val=\"", highlightElementStart);
                if (valStart < 0) {
                    // Try without namespace prefix
                    valStart = rprXml.indexOf("val=\"", highlightElementStart);
                    if (valStart > 0) {
                        valStart += 5; // Skip "val=\""
                    }
                } else {
                    valStart += 7; // Skip "w:val=\""
                }
                
                if (valStart > highlightElementStart) {
                    int valEnd = rprXml.indexOf("\"", valStart);
                    if (valEnd > valStart) {
                        return rprXml.substring(valStart, valEnd);
                    }
                }
            }
        } catch (Exception e) {
            // Return null if extraction fails
        }
        return null;
    }
}


