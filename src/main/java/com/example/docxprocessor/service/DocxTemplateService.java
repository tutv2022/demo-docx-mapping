package com.example.docxprocessor.service;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class DocxTemplateService {

    // New syntax patterns: {{$variable}}, {{#loop:$array}}, {{#/loop}}, etc. (keeps $ but removed json.)
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\{\\{\\$([^}]+)\\}\\}");
    private static final Pattern CHECKBOX_PATTERN = Pattern.compile("\\{\\{#checkbox:\\s*\\$([^:]+):([^}]+)\\}\\}");
    private static final Pattern RADIO_PATTERN = Pattern.compile("\\{\\{#radio:\\s*\\$([^:]+):([^}]+)\\}\\}");
    // Loop patterns - new syntax: {{#loop:$arrayName}} and {{#/loop}}
    private static final Pattern LOOP_START_PATTERN = Pattern.compile("\\{\\{#loop:\\s*\\$([^}]+)\\}\\}");
    private static final Pattern LOOP_END_PATTERN = Pattern.compile("\\{\\{#/loop\\}\\}");
    // Pattern for loop item access: {{$arrayName[i].property}} or {{$arrayName[1].property}}
    // Supports both [i] (current index) and [number] (specific index)
    // Allows optional whitespace and handles various property names
    private static final Pattern LOOP_ITEM_PATTERN = Pattern.compile("\\{\\{\\s*\\$([^\\[\\}]+)\\[([^\\]]+)\\]\\.([^\\}]+)\\s*\\}\\}");

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
                processAllPlaceholders(paragraph, data, null, -1, data, null);
            }

            // Process tables
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            processAllPlaceholders(paragraph, data, null, -1, data, null);
                        }
                    }
                }
            }

            // Process headers
            for (XWPFHeader header : document.getHeaderList()) {
                for (XWPFParagraph paragraph : header.getParagraphs()) {
                    processAllPlaceholders(paragraph, data, null, -1, data, null);
                }
            }

            // Process footers
            for (XWPFFooter footer : document.getFooterList()) {
                for (XWPFParagraph paragraph : footer.getParagraphs()) {
                    processAllPlaceholders(paragraph, data, null, -1, data, null);
                }
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    /**
     * Helper method to extract variable name from $variableName format
     * @param variablePath The variable path like "$customerName" or "$orderItems"
     * @return The variable name without $ prefix
     */
    private String extractVariableName(String variablePath) {
        if (variablePath == null) {
            return null;
        }
        // Remove $ prefix if present
        String trimmed = variablePath.trim();
        if (trimmed.startsWith("$")) {
            return trimmed.substring(1);
        }
        return trimmed;
    }

    /**
     * Processes loop blocks in the document
     * New syntax: {{#loop:$arrayName}} ... content ... {{#/loop}}
     * Within loops, use {{$arrayName[i].property}} to access item properties
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
        
        String loopVariablePath = startMatcher.group(1); // e.g., "orderItems"
        String loopVariable = extractVariableName(loopVariablePath);
        
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
        
        // Use indexed loop to track current index for [i] syntax
        for (int itemIndex = 0; itemIndex < items.size(); itemIndex++) {
            Map<String, Object> item = items.get(itemIndex);
            
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
                
                // Process placeholders with item data, loop context, and current index
                // Pass both the item and the original data map so we can access the list by index
                // This must work correctly - if paragraph has no runs, it means it's empty and we skip it
                if (!newPara.getRuns().isEmpty()) {
                    processAllPlaceholders(newPara, item, loopVariable, itemIndex, data, items);
                } else {
                    // If no runs, check if there's text in the CT element that needs processing
                    String paraText = getParagraphText(newPara);
                    if (paraText != null && !paraText.trim().isEmpty()) {
                        // There's text but no runs - this shouldn't happen with a cloned paragraph
                        // But if it does, create a run and process it
                        XWPFRun run = newPara.createRun();
                        run.setText(paraText);
                        processAllPlaceholders(newPara, item, loopVariable, itemIndex, data, items);
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
     * Processes radio buttons, checkboxes, and regular placeholders while preserving formatting
     * only for the placeholder variables themselves, not the entire paragraph.
     * New syntax: {{$variable}}, {{#checkbox:$var:value}}, {{#radio:$var:value}}
     * 
     * @param paragraph The paragraph to process
     * @param data The data map (current item when in loop, or full data when not in loop)
     * @param loopArrayName Optional loop array name for handling [i] syntax (null if not in loop)
     * @param currentIndex Current index in the loop (for [i] syntax), -1 if not in loop
     * @param originalData Original data map containing the list (for loop item access)
     * @param itemsList The list of items (for loop item access by index)
     */
    private void processAllPlaceholders(XWPFParagraph paragraph, Map<String, Object> data, String loopArrayName, int currentIndex, Map<String, Object> originalData, List<Map<String, Object>> itemsList) {
        if (paragraph.getRuns().isEmpty()) {
            return;
        }

        // Collect all text from runs to check for placeholders
        String paragraphText = getParagraphText(paragraph);
        
        // Check if paragraph has any placeholders
        Matcher radioMatcher = RADIO_PATTERN.matcher(paragraphText);
        Matcher checkboxMatcher = CHECKBOX_PATTERN.matcher(paragraphText);
        Matcher regularMatcher = PLACEHOLDER_PATTERN.matcher(paragraphText);
        Matcher loopItemMatcher = LOOP_ITEM_PATTERN.matcher(paragraphText);
        
        boolean hasRadios = radioMatcher.find();
        boolean hasCheckboxes = checkboxMatcher.find();
        boolean hasRegularPlaceholders = regularMatcher.find();
        boolean hasLoopItems = loopItemMatcher.find();
        
        if (!hasRadios && !hasCheckboxes && !hasRegularPlaceholders && !hasLoopItems) {
            return; // No placeholders in this paragraph
        }

        // Build replacement map for all placeholders
        Map<String, String> allReplacements = new HashMap<>();

        // Step 1: Build replacement map for loop item placeholders ({{$arrayName[i].property}} or {{$arrayName[1].property}})
        // Supports both [i] (current index) and [number] (specific index)
        if (hasLoopItems && originalData != null) {
            loopItemMatcher.reset();
            while (loopItemMatcher.find()) {
                String arrayName = loopItemMatcher.group(1).trim();
                String indexStr = loopItemMatcher.group(2).trim(); // Can be "i" or a number like "1", "2", etc.
                String property = loopItemMatcher.group(3).trim();
                String fullPlaceholder = loopItemMatcher.group(0); // e.g., {{$orderItems[i].unitPrice}} or {{$orderItems[1].itemSKU}}
                
                // Determine which index to use
                int targetIndex = -1;
                if (indexStr.equalsIgnoreCase("i")) {
                    // Use current loop index
                    if (loopArrayName != null && arrayName.equals(loopArrayName) && currentIndex >= 0) {
                        targetIndex = currentIndex;
                    } else {
                        // Not in a loop context or array name doesn't match, skip this placeholder
                        continue;
                    }
                } else {
                    // Try to parse as a number (specific index)
                    try {
                        targetIndex = Integer.parseInt(indexStr);
                    } catch (NumberFormatException e) {
                        // Invalid index, skip this placeholder
                        continue;
                    }
                }
                
                // Extract value from original data structure using the path: arrayName[index].property
                Object propertyValue = null;
                
                // Get the list from original data
                Object listObj = originalData.get(arrayName);
                if (listObj == null) {
                    // List not found, skip
                    continue;
                } else if (!(listObj instanceof List)) {
                    // Not a list, skip
                    continue;
                } else {
                    @SuppressWarnings("unchecked")
                    List<Object> items = (List<Object>) listObj;
                    
                    if (targetIndex < 0 || targetIndex >= items.size()) {
                        // Index out of bounds, use empty value
                        propertyValue = null;
                    } else {
                        Object itemObj = items.get(targetIndex);
                        
                        if (itemObj == null) {
                            propertyValue = null;
                        } else if (!(itemObj instanceof Map)) {
                            propertyValue = null;
                        } else {
                            @SuppressWarnings("unchecked")
                            Map<String, Object> item = (Map<String, Object>) itemObj;
                            
                            // Try exact match first
                            if (item.containsKey(property)) {
                                propertyValue = item.get(property);
                            } else {
                                // Try case-insensitive lookup
                                for (Map.Entry<String, Object> entry : item.entrySet()) {
                                    if (entry.getKey().equalsIgnoreCase(property)) {
                                        propertyValue = entry.getValue();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                
                String value = propertyValue != null ? propertyValue.toString() : "";
                // Use the original placeholder as the key for replacement
                allReplacements.put(fullPlaceholder, value);
            }
        }
        
        // Debug: show all replacements
        if (!allReplacements.isEmpty()) {
            System.out.println("DEBUG: Total replacements: " + allReplacements.size() + " - " + allReplacements);
        }

        // Step 2: Build replacement map for radio button placeholders
        if (hasRadios) {
            radioMatcher.reset();
            while (radioMatcher.find()) {
                String radioGroupPath = radioMatcher.group(1);
                String radioValue = radioMatcher.group(2);
                String fullPlaceholder = radioMatcher.group(0);
                
                String radioGroup = extractVariableName(radioGroupPath);
                Object groupValueObj = data.get(radioGroup);
                String groupValue = groupValueObj != null ? groupValueObj.toString() : null;
                boolean isSelected = radioValue.equalsIgnoreCase(groupValue);
                
                String radioSymbol = isSelected ? "◉" : "○";
                allReplacements.put(fullPlaceholder, radioSymbol);
            }
        }

        // Step 3: Build replacement map for checkbox placeholders
        if (hasCheckboxes) {
            checkboxMatcher.reset();
            while (checkboxMatcher.find()) {
                String checkboxGroupPath = checkboxMatcher.group(1);
                String checkboxValue = checkboxMatcher.group(2);
                String fullPlaceholder = checkboxMatcher.group(0);
                
                String checkboxGroup = extractVariableName(checkboxGroupPath);
                Object groupValueObj = data.get(checkboxGroup);
                String groupValue = groupValueObj != null ? groupValueObj.toString() : null;
                boolean isChecked = checkboxValue.equalsIgnoreCase(groupValue);
                
                String checkboxSymbol = isChecked ? "☒" : "☐";
                allReplacements.put(fullPlaceholder, checkboxSymbol);
            }
        }

        // Step 4: Build replacement map for regular placeholders ({{$variableName}})
        // IMPORTANT: Skip placeholders that were already processed in Step 1 (loop items)
        // Loop item placeholders like {{$orderItems[i].property}} or {{$orderItems[1].property}} should not be processed here
        if (hasRegularPlaceholders) {
            regularMatcher.reset();
            while (regularMatcher.find()) {
                String variablePath = regularMatcher.group(1);
                String fullPlaceholder = regularMatcher.group(0);

                // Skip if this placeholder contains [i] or [number] - it's a loop item placeholder
                // Loop item placeholders are already processed in Step 1
                if (fullPlaceholder.matches(".*\\[([i]|\\d+)\\].*")) {
                    continue;
                }
                
                // Skip if this placeholder was already added in Step 1
                if (allReplacements.containsKey(fullPlaceholder)) {
                    continue;
                }

                String variableName = extractVariableName(variablePath);
                Object valueObj = data.get(variableName);
                String value = valueObj != null ? valueObj.toString() : "";
                allReplacements.put(fullPlaceholder, value);
            }
        }

        // Step 5: Process runs individually to preserve formatting only for placeholders
        if (!allReplacements.isEmpty()) {
            // Debug: show all replacements
            // System.out.println("DEBUG: All replacements: " + allReplacements);
            processRunsWithPlaceholderFormatting(paragraph, allReplacements, currentIndex);
        }
    }

    /**
     * Processes runs individually, replacing placeholders while preserving their original formatting
     * This ensures only placeholder text gets its formatting preserved, not the entire paragraph
     * Handles placeholders that may span multiple runs
     */
    private void processRunsWithPlaceholderFormatting(XWPFParagraph paragraph, Map<String, String> replacements, int currentIndex) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs.isEmpty()) {
            return;
        }

        // Get full paragraph text to find all placeholders
        String fullText = getParagraphText(paragraph);
        if (fullText == null || fullText.isEmpty()) {
            return;
        }

        // Process placeholders that exist in the full text
        // This handles both single-run and multi-run placeholders
        // Process from the end of the text backwards to avoid index shifting
        List<PlaceholderPosition> placeholderPositions = new ArrayList<>();
        
        // First, find all loop item placeholders in the text using pattern matching
        // This ensures we catch them even if whitespace differs
        // Supports both [i] and [number] syntax
        Matcher loopItemMatcher = LOOP_ITEM_PATTERN.matcher(fullText);
        Map<String, String> loopItemMatches = new HashMap<>(); // Map from matched text to replacement value
        while (loopItemMatcher.find()) {
            String matchedText = loopItemMatcher.group(0);
            String arrayName = loopItemMatcher.group(1).trim();
            String indexStr = loopItemMatcher.group(2).trim(); // Can be "i" or a number
            String property = loopItemMatcher.group(3).trim();
            
            // Build the key that should be in replacements map (with the actual index from match)
            String placeholderKey = "{{$" + arrayName + "[" + indexStr + "]." + property + "}}";
            
            // Find matching replacement (try exact match first)
            String replacementValue = replacements.get(placeholderKey);
            if (replacementValue == null) {
                // Try to find by matching array name, index, and property
                for (Map.Entry<String, String> entry : replacements.entrySet()) {
                    String key = entry.getKey();
                    if (key.contains("[" + indexStr + "]") && key.contains(arrayName) && key.contains(property)) {
                        replacementValue = entry.getValue();
                        break;
                    }
                }
            }
            
            if (replacementValue != null) {
                loopItemMatches.put(matchedText, replacementValue);
                placeholderPositions.add(new PlaceholderPosition(loopItemMatcher.start(), matchedText, replacementValue));
            }
        }
        
        // Then find other placeholders (non-loop items) using exact matching
        for (Map.Entry<String, String> replacement : replacements.entrySet()) {
            String placeholder = replacement.getKey();
            String replacementValue = replacement.getValue();
            
            // Skip if already processed as loop item
            if (loopItemMatches.containsKey(placeholder)) {
                continue;
            }
            
            // Try exact match
            int searchStart = 0;
            while (true) {
                int pos = fullText.indexOf(placeholder, searchStart);
                if (pos < 0) {
                    break;
                }
                placeholderPositions.add(new PlaceholderPosition(pos, placeholder, replacementValue));
                searchStart = pos + 1;
            }
        }
        
        // Sort by position descending so we process from end to start
        placeholderPositions.sort((a, b) -> Integer.compare(b.position, a.position));
        
        // Process each placeholder
        for (PlaceholderPosition pp : placeholderPositions) {
            // Refresh runs and fullText as they may have changed
            runs = paragraph.getRuns();
            fullText = getParagraphText(paragraph);
            
            // Check if placeholder still exists (might have been replaced already)
            if (!fullText.contains(pp.placeholder)) {
                continue;
            }
            
            // Find the position again (might have shifted)
            int placeholderPos = fullText.indexOf(pp.placeholder);
            if (placeholderPos < 0) {
                continue;
            }
            
            // Find which run(s) contain this placeholder
            RunRange runRange = findRunRangeForTextPosition(runs, placeholderPos, placeholderPos + pp.placeholder.length());
            if (runRange == null) {
                // Debug: placeholder not found in runs
                // System.out.println("DEBUG processRuns: Could not find run range for placeholder: " + pp.placeholder + " at position: " + placeholderPos);
                continue;
            }
            
            // Debug: found run range
            // System.out.println("DEBUG processRuns: Found run range for " + pp.placeholder + ": startRun=" + runRange.startRunIndex + ", endRun=" + runRange.endRunIndex);
            
            // Extract formatting from the first run containing the placeholder
            RunFormatting formatting = extractRunFormatting(runs.get(runRange.startRunIndex));
            
            if (runRange.startRunIndex == runRange.endRunIndex) {
                // Placeholder is in a single run
                XWPFRun targetRun = runs.get(runRange.startRunIndex);
                String runText = getRunText(targetRun);
                if (runText != null && runText.contains(pp.placeholder)) {
                    String newRunText = runText.replace(pp.placeholder, pp.replacementValue);
                    int runIdx = runRange.startRunIndex;
                    paragraph.removeRun(runIdx);
                    XWPFRun newRun = paragraph.insertNewRun(runIdx);
                    newRun.setText(newRunText);
                    applyRunFormatting(formatting, newRun);
                }
            } else {
                // Placeholder spans multiple runs - merge them
                StringBuilder combinedText = new StringBuilder();
                for (int i = runRange.startRunIndex; i <= runRange.endRunIndex; i++) {
                    String runText = getRunText(runs.get(i));
                    if (runText != null) {
                        combinedText.append(runText);
                    }
                }
                
                String combined = combinedText.toString();
                if (combined.contains(pp.placeholder)) {
                    String newText = combined.replace(pp.placeholder, pp.replacementValue);
                    
                    // Remove all runs in the range (from end to start)
                    for (int i = runRange.endRunIndex; i >= runRange.startRunIndex; i--) {
                        paragraph.removeRun(i);
                    }
                    
                    // Insert new run with replaced text and preserved formatting
                    XWPFRun newRun = paragraph.insertNewRun(runRange.startRunIndex);
                    newRun.setText(newText);
                    applyRunFormatting(formatting, newRun);
                }
            }
        }
        
        // Refresh runs and fullText after placeholder replacements
        runs = paragraph.getRuns();
        fullText = getParagraphText(paragraph);
        
        
        // Also handle standalone [i] replacement if in loop context
        // But ONLY if it's not part of a placeholder pattern
        if (currentIndex >= 0) {
            runs = paragraph.getRuns(); // Refresh runs list
            fullText = getParagraphText(paragraph); // Get updated text
            
            // Check if there are any remaining placeholders with [i] that should be replaced
            // If there are, don't replace standalone [i] as it might break the placeholder
            boolean hasUnreplacedPlaceholders = false;
            for (String placeholder : replacements.keySet()) {
                if (placeholder.contains("[i]") && fullText.contains(placeholder)) {
                    hasUnreplacedPlaceholders = true;
                    break;
                }
            }
            
            // Only replace standalone [i] if there are no unreplaced placeholders
            if (!hasUnreplacedPlaceholders) {
                for (int i = runs.size() - 1; i >= 0; i--) {
                    XWPFRun run = runs.get(i);
                    String runText = getRunText(run);
                    if (runText != null && runText.contains("[i]")) {
                        // Make sure [i] is not part of a placeholder pattern
                        int iPos = runText.indexOf("[i]");
                        if (iPos > 0) {
                            String beforeI = runText.substring(Math.max(0, iPos - 30), iPos);
                            // If [i] is preceded by {{, it's part of a placeholder pattern
                            if (!beforeI.contains("{{")) {
                                RunFormatting formatting = extractRunFormatting(run);
                                String newRunText = runText.replace("[i]", String.valueOf(currentIndex));
                                paragraph.removeRun(i);
                                XWPFRun newRun = paragraph.insertNewRun(i);
                                newRun.setText(newRunText);
                                applyRunFormatting(formatting, newRun);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Finds which run(s) contain a text position range
     */
    private RunRange findRunRangeForTextPosition(List<XWPFRun> runs, int startPos, int endPos) {
        int currentPos = 0;
        int startRunIndex = -1;
        int endRunIndex = -1;
        XWPFRun startRun = null;
        
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String runText = getRunText(run);
            if (runText == null || runText.isEmpty()) {
                continue;
            }
            
            int runStart = currentPos;
            int runEnd = currentPos + runText.length();
            
            // Check if this run contains the start position (inclusive start, exclusive end)
            if (startRunIndex == -1 && startPos >= runStart && startPos < runEnd) {
                startRunIndex = i;
                startRun = run;
            }
            
            // Check if this run contains the end position (exclusive end)
            // endPos is exclusive, so we check if it's > runStart and <= runEnd
            if (endPos > runStart && endPos <= runEnd) {
                endRunIndex = i;
                // Don't break yet - we need to make sure we found the start too
                if (startRunIndex >= 0) {
                    break;
                }
            }
            
            currentPos = runEnd;
        }
        
        // If we found both start and end, return the range
        // If we only found start, assume it's in a single run
        if (startRunIndex >= 0) {
            if (endRunIndex < 0) {
                endRunIndex = startRunIndex; // Single run
            }
            return new RunRange(startRunIndex, endRunIndex, startRun);
        }
        
        return null;
    }

    /**
     * Helper class to track run range for a placeholder
     */
    private static class RunRange {
        int startRunIndex;
        int endRunIndex;
        XWPFRun startRun;
        
        RunRange(int startRunIndex, int endRunIndex, XWPFRun startRun) {
            this.startRunIndex = startRunIndex;
            this.endRunIndex = endRunIndex;
            this.startRun = startRun;
        }
    }

    /**
     * Helper class to track placeholder position and replacement
     */
    private static class PlaceholderPosition {
        int position;
        String placeholder;
        String replacementValue;
        
        PlaceholderPosition(int position, String placeholder, String replacementValue) {
            this.position = position;
            this.placeholder = placeholder;
            this.replacementValue = replacementValue;
        }
    }


    /**
     * Safely gets text from a run
     */
    private String getRunText(XWPFRun run) {
        if (run == null) {
            return null;
        }
        try {
            return run.getText(0);
        } catch (IndexOutOfBoundsException e) {
            return null;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Overloaded method for backward compatibility (non-loop context)
     */
    private void processAllPlaceholders(XWPFParagraph paragraph, Map<String, Object> data) {
        processAllPlaceholders(paragraph, data, null, -1, data, null);
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


