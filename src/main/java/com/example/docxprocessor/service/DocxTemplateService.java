package com.example.docxprocessor.service;

import com.jayway.jsonpath.DocumentContext;
import com.jayway.jsonpath.JsonPath;
import com.jayway.jsonpath.PathNotFoundException;
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
     * Processes a DOCX template file using JSON-Path expressions to extract values from JSON
     * Placeholders like {{$orderId}} or {{$customer.name}} will use JSON-Path to extract values directly
     * 
     * @param templateInputStream Input stream of the template DOCX file
     * @param jsonString JSON string to extract values from using JSON-Path
     * @return Byte array of the processed DOCX file
     * @throws IOException if file processing fails
     */
    public byte[] processTemplateWithJsonPath(InputStream templateInputStream, String jsonString) throws IOException {
        DocumentContext jsonContext = JsonPath.parse(jsonString);
        return processTemplateWithJsonPath(templateInputStream, jsonContext);
    }

    /**
     * Processes a DOCX template file using JSON-Path DocumentContext
     * Uses JSON-Path expressions directly when replacing placeholders
     */
    public byte[] processTemplateWithJsonPath(InputStream templateInputStream, DocumentContext jsonContext) throws IOException {
        try (XWPFDocument document = new XWPFDocument(templateInputStream)) {
            
            // Step 1: Process loops first (they can contain placeholders)
            processLoopsWithJsonPath(document, jsonContext);
            
            // Step 2: Process table row loops BEFORE processing regular placeholders
            for (XWPFTable table : document.getTables()) {
                processTableLoopsWithJsonPath(table, jsonContext);
            }
            
            // Step 3: Process paragraphs - unified processing for checkboxes and regular placeholders
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                processAllPlaceholdersWithJsonPath(paragraph, jsonContext, null, -1, jsonContext, null);
            }

            // Step 4: Process tables - regular placeholders (after loops are processed)
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            processAllPlaceholdersWithJsonPath(paragraph, jsonContext, null, -1, jsonContext, null);
                        }
                    }
                }
            }

            // Process headers
            for (XWPFHeader header : document.getHeaderList()) {
                for (XWPFParagraph paragraph : header.getParagraphs()) {
                    processAllPlaceholdersWithJsonPath(paragraph, jsonContext, null, -1, jsonContext, null);
                }
            }

            // Process footers
            for (XWPFFooter footer : document.getFooterList()) {
                for (XWPFParagraph paragraph : footer.getParagraphs()) {
                    processAllPlaceholdersWithJsonPath(paragraph, jsonContext, null, -1, jsonContext, null);
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
     * Removes a placeholder pattern from a paragraph while preserving formatting
     * Preserves other text in the paragraph and maintains original formatting (bold, italic, etc.)
     * Removes ALL occurrences of the pattern, not just the first one
     */
    private void removePlaceholderFromParagraph(XWPFParagraph paragraph, Pattern pattern) {
        if (paragraph == null) {
            return;
        }
        
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs.isEmpty()) {
            return;
        }
        
        String fullText = getParagraphText(paragraph);
        if (fullText == null || fullText.isEmpty()) {
            return;
        }
        
        // Check if the pattern exists in the text
        Matcher matcher = pattern.matcher(fullText);
        if (!matcher.find()) {
            return; // Pattern not found, nothing to remove
        }
        
        // Find all occurrences of the pattern
        List<int[]> matches = new ArrayList<>();
        matcher.reset();
        while (matcher.find()) {
            matches.add(new int[]{matcher.start(), matcher.end()});
        }
        
        // Process from end to start to avoid index shifting
        matches.sort((a, b) -> Integer.compare(b[0], a[0]));
        
        for (int[] match : matches) {
            int startPos = match[0];
            int endPos = match[1];
            
            // Find which run(s) contain this marker
            RunRange runRange = findRunRangeForTextPosition(runs, startPos, endPos);
            if (runRange == null) {
                continue;
            }
            
            // Calculate positions within runs
            int currentPos = 0;
            int startPosInStartRun = -1;
            int endPosInEndRun = -1;
            
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = getRunText(run);
                if (runText == null || runText.isEmpty()) {
                    continue;
                }
                
                int runStart = currentPos;
                int runEnd = currentPos + runText.length();
                
                if (i == runRange.startRunIndex && startPos >= runStart && startPos < runEnd) {
                    startPosInStartRun = startPos - runStart;
                }
                
                if (i == runRange.endRunIndex && endPos > runStart && endPos <= runEnd) {
                    endPosInEndRun = endPos - runStart;
                }
                
                currentPos = runEnd;
            }
            
            // Remove the marker from the run(s) while preserving formatting
            if (runRange.startRunIndex == runRange.endRunIndex) {
                // Marker is in a single run
                XWPFRun run = runs.get(runRange.startRunIndex);
                String runText = getRunText(run);
                if (runText != null && startPosInStartRun >= 0 && endPosInEndRun >= 0) {
                    // Remove the marker from this run
                    String beforeMarker = runText.substring(0, startPosInStartRun);
                    String afterMarker = runText.substring(endPosInEndRun);
                    String newRunText = beforeMarker + afterMarker;
                    
                    // Extract formatting before modifying
                    RunFormatting formatting = extractRunFormatting(run);
                    
                    // Update the run text (preserves formatting)
                    run.setText(newRunText, 0);
                    
                    // Reapply formatting
                    applyRunFormatting(formatting, run);
                }
            } else {
                // Marker spans multiple runs
                // First run: keep text before marker
                if (runRange.startRunIndex >= 0 && runRange.startRunIndex < runs.size()) {
                    XWPFRun firstRun = runs.get(runRange.startRunIndex);
                    String firstRunText = getRunText(firstRun);
                    if (firstRunText != null && startPosInStartRun >= 0) {
                        RunFormatting formatting = extractRunFormatting(firstRun);
                        String beforeMarker = firstRunText.substring(0, startPosInStartRun);
                        firstRun.setText(beforeMarker, 0);
                        applyRunFormatting(formatting, firstRun);
                    } else {
                        firstRun.setText("", 0);
                    }
                }
                
                // Last run: keep text after marker
                if (runRange.endRunIndex >= 0 && runRange.endRunIndex < runs.size()) {
                    XWPFRun lastRun = runs.get(runRange.endRunIndex);
                    String lastRunText = getRunText(lastRun);
                    if (lastRunText != null && endPosInEndRun >= 0 && endPosInEndRun < lastRunText.length()) {
                        RunFormatting formatting = extractRunFormatting(lastRun);
                        String afterMarker = lastRunText.substring(endPosInEndRun);
                        lastRun.setText(afterMarker, 0);
                        applyRunFormatting(formatting, lastRun);
                    } else {
                        lastRun.setText("", 0);
                    }
                }
                
                // Remove middle runs (if any)
                for (int i = runRange.endRunIndex - 1; i > runRange.startRunIndex; i--) {
                    try {
                        paragraph.removeRun(i);
                    } catch (Exception e) {
                        // If removal fails, just clear the text
                        runs.get(i).setText("", 0);
                    }
                }
            }
        }
        
        // Clean up empty runs
        cleanupEmptyRuns(paragraph);
    }
    
    /**
     * Removes empty runs from a paragraph
     */
    private void cleanupEmptyRuns(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = runs.size() - 1; i >= 0; i--) {
            try {
                XWPFRun run = runs.get(i);
                String text = run.getText(0);
                if (text == null || text.trim().isEmpty()) {
                    paragraph.removeRun(i);
                }
            } catch (Exception e) {
                // Skip if we can't check or remove
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
            
            // Build the key that should be in replacements map
            // The key in replacements map has [i] or [number], we need to match it
            String placeholderKey = "{{$" + arrayName + "[" + indexStr + "]." + property + "}}";
            
            // Normalize the key (remove extra whitespace) to match
            String normalizedKey = placeholderKey.replaceAll("\\s+", "");
            
            // Find matching replacement (try exact match first)
            String replacementValue = replacements.get(placeholderKey);
            if (replacementValue == null) {
                // Try normalized key
                replacementValue = replacements.get(normalizedKey);
            }
            if (replacementValue == null) {
                // Try to find by matching array name, index, and property (case-insensitive)
                for (Map.Entry<String, String> entry : replacements.entrySet()) {
                    String key = entry.getKey();
                    String normalizedEntryKey = key.replaceAll("\\s+", "");
                    if (normalizedEntryKey.equalsIgnoreCase(normalizedKey)) {
                        replacementValue = entry.getValue();
                        break;
                    }
                    // Also try partial match
                    if (key.contains("[" + indexStr + "]") && 
                        key.toLowerCase().contains(arrayName.toLowerCase()) && 
                        key.contains(property)) {
                        replacementValue = entry.getValue();
                        break;
                    }
                }
            }
            
            if (replacementValue != null) {
                loopItemMatches.put(matchedText, replacementValue);
                placeholderPositions.add(new PlaceholderPosition(loopItemMatcher.start(), matchedText, replacementValue));
            } else {
                // If no replacement found, replace with empty string
                loopItemMatches.put(matchedText, "");
                placeholderPositions.add(new PlaceholderPosition(loopItemMatcher.start(), matchedText, ""));
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
            
            // Replace the placeholder with the replacement value (empty string if path not found)
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

    // ==================== JSON-Path Implementation Methods ====================

    /**
     * Processes loop blocks in the document using JSON-Path
     * New syntax: {{#loop:$arrayName}} ... content ... {{#/loop}}
     * Within loops, use {{$arrayName[i].property}} to access item properties
     */
    private void processLoopsWithJsonPath(XWPFDocument document, DocumentContext jsonContext) {
        List<XWPFParagraph> paragraphs = new ArrayList<>(document.getParagraphs());
        
        if (paragraphs.isEmpty()) {
            return;
        }
        
        int maxIterations = 100;
        int iteration = 0;
        
        while (iteration < maxIterations) {
            iteration++;
            boolean foundLoop = false;
            
            for (int i = paragraphs.size() - 1; i >= 0; i--) {
                XWPFParagraph paragraph = paragraphs.get(i);
                if (paragraph == null) {
                    continue;
                }
                
                String text = getParagraphText(paragraph);
                Matcher endMatcher = LOOP_END_PATTERN.matcher(text);
                if (endMatcher.find()) {
                    int startIndex = findLoopStartMarker(paragraphs, i);
                    if (startIndex != -1 && startIndex < i) {
                        processLoopBlockWithJsonPath(document, paragraphs, startIndex, i, jsonContext);
                        foundLoop = true;
                        paragraphs = new ArrayList<>(document.getParagraphs());
                        break;
                    }
                }
            }
            
            if (!foundLoop) {
                break;
            }
        }
    }

    /**
     * Processes a single loop block using JSON-Path
     */
    @SuppressWarnings("unchecked")
    private void processLoopBlockWithJsonPath(XWPFDocument document, List<XWPFParagraph> paragraphs, 
                                              int startIndex, int endIndex, DocumentContext jsonContext) {
        XWPFParagraph startPara = paragraphs.get(startIndex);
        String startText = getParagraphText(startPara);
        
        Matcher startMatcher = LOOP_START_PATTERN.matcher(startText);
        if (!startMatcher.find()) {
            return;
        }
        
        String loopVariablePath = startMatcher.group(1);
        String loopVariable = extractVariableName(loopVariablePath);
        
        // Use JSON-Path to get the list
        String jsonPathExpr = "$." + loopVariable;
        Object listObj;
        try {
            listObj = jsonContext.read(jsonPathExpr);
        } catch (PathNotFoundException e) {
            removeLoopMarkers(paragraphs, startIndex, endIndex);
            return;
        }
        
        if (!(listObj instanceof List)) {
            removeLoopMarkers(paragraphs, startIndex, endIndex);
            return;
        }
        
        List<Object> items = (List<Object>) listObj;
        if (items.isEmpty()) {
            removeLoopBlock(paragraphs, startIndex, endIndex);
            return;
        }
        
        // Extract content between markers
        List<XWPFParagraph> loopContent = new ArrayList<>();
        for (int i = startIndex + 1; i < endIndex; i++) {
            loopContent.add(paragraphs.get(i));
        }
        
        List<String> loopContentXml = new ArrayList<>();
        for (XWPFParagraph para : loopContent) {
            String paraXml = para.getCTP().xmlText();
            loopContentXml.add(paraXml);
        }
        
        removePlaceholderFromParagraph(startPara, LOOP_START_PATTERN);
        XWPFParagraph endPara = paragraphs.get(endIndex);
        removePlaceholderFromParagraph(endPara, LOOP_END_PATTERN);
        
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody body = document.getDocument().getBody();
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP startParaCT = startPara.getCTP();
        
        int startParaIndex = -1;
        String startParaXml = startParaCT.xmlText();
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i).xmlText().equals(startParaXml)) {
                startParaIndex = i;
                break;
            }
        }
        
        if (startParaIndex == -1) {
            return;
        }
        
        int endParaIndex = -1;
        String endParaXml = endPara.getCTP().xmlText();
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i).xmlText().equals(endParaXml)) {
                endParaIndex = i;
                break;
            }
        }
        
        if (endParaIndex == -1 || endParaIndex <= startParaIndex) {
            return;
        }
        
        body.removeP(endParaIndex);
        for (int i = endParaIndex - 1; i > startParaIndex; i--) {
            body.removeP(i);
        }
        
        int currentInsertPos = startParaIndex + 1;
        
        for (int itemIndex = 0; itemIndex < items.size(); itemIndex++) {
            Object item = items.get(itemIndex);
            
            for (String paraXml : loopContentXml) {
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP clonedCT;
                try {
                    clonedCT = org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP.Factory.parse(paraXml);
                } catch (org.apache.xmlbeans.XmlException e) {
                    continue;
                }
                
                body.insertNewP(currentInsertPos);
                body.setPArray(currentInsertPos, clonedCT);
                
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP newParaCT = body.getPArray(currentInsertPos);
                XWPFParagraph newPara = new XWPFParagraph(newParaCT, document);
                
                if (!newPara.getRuns().isEmpty()) {
                    // For loop items, we use the original context and access items by index
                    // The placeholder {{$arrayName[i].property}} will be resolved using the original context
                    processAllPlaceholdersWithJsonPath(newPara, jsonContext, loopVariable, itemIndex, jsonContext, items);
                } else {
                    String paraText = getParagraphText(newPara);
                    if (paraText != null && !paraText.trim().isEmpty()) {
                        XWPFRun run = newPara.createRun();
                        run.setText(paraText);
                        processAllPlaceholdersWithJsonPath(newPara, jsonContext, loopVariable, itemIndex, jsonContext, items);
                    }
                }
                
                currentInsertPos++;
            }
        }
    }

    /**
     * Process placeholders using JSON-Path expressions directly
     * Extracts values directly from JSON using JSON-Path when replacing placeholders
     */
    private void processAllPlaceholdersWithJsonPath(XWPFParagraph paragraph, DocumentContext jsonContext, 
                                                     String loopArrayName, int currentIndex, 
                                                     DocumentContext originalJsonContext, 
                                                     List<Object> itemsList) {
        if (paragraph.getRuns().isEmpty()) {
            return;
        }

        String paragraphText = getParagraphText(paragraph);
        
        // Debug: Log paragraph text and currentIndex
        System.out.println("DEBUG processAllPlaceholdersWithJsonPath: paragraphText='" + paragraphText + "', currentIndex=" + currentIndex + ", loopArrayName=" + loopArrayName);
        
        if (paragraphText == null || paragraphText.trim().isEmpty()) {
            System.out.println("DEBUG: Paragraph text is null or empty, returning");
            return;
        }
        
        Matcher radioMatcher = RADIO_PATTERN.matcher(paragraphText);
        Matcher checkboxMatcher = CHECKBOX_PATTERN.matcher(paragraphText);
        Matcher regularMatcher = PLACEHOLDER_PATTERN.matcher(paragraphText);
        Matcher loopItemMatcher = LOOP_ITEM_PATTERN.matcher(paragraphText);
        
        boolean hasRadios = radioMatcher.find();
        boolean hasCheckboxes = checkboxMatcher.find();
        boolean hasRegularPlaceholders = regularMatcher.find();
        boolean hasLoopItems = loopItemMatcher.find();
        
        System.out.println("DEBUG: hasRadios=" + hasRadios + ", hasCheckboxes=" + hasCheckboxes + ", hasRegularPlaceholders=" + hasRegularPlaceholders + ", hasLoopItems=" + hasLoopItems);
        
        if (!hasRadios && !hasCheckboxes && !hasRegularPlaceholders && !hasLoopItems) {
            System.out.println("DEBUG: No placeholders found, returning");
            return;
        }

        Map<String, String> allReplacements = new HashMap<>();

        // Step 1: Process loop item placeholders using JSON-Path
        // Replace [i] with actual numeric index
        if (hasLoopItems) {
            System.out.println("DEBUG: Processing loop items, originalJsonContext is " + (originalJsonContext != null ? "not null" : "null"));
            loopItemMatcher.reset();
            int matchCount = 0;
            while (loopItemMatcher.find()) {
                matchCount++;
                String arrayName = loopItemMatcher.group(1).trim();
                String indexStr = loopItemMatcher.group(2).trim();
                String property = loopItemMatcher.group(3).trim();
                String fullPlaceholder = loopItemMatcher.group(0);
                
                System.out.println("DEBUG: Found loop item placeholder: " + fullPlaceholder + ", arrayName=" + arrayName + ", indexStr=" + indexStr + ", property=" + property);
                
                // If originalJsonContext is null, replace with empty string
                if (originalJsonContext == null) {
                    System.out.println("WARN: originalJsonContext is null for placeholder: " + fullPlaceholder + " - Replacing with empty string");
                    allReplacements.put(fullPlaceholder, "");
                    continue;
                }
                
                int targetIndex = -1;
                if (indexStr.equalsIgnoreCase("i")) {
                    System.out.println("DEBUG: indexStr is 'i', currentIndex=" + currentIndex);
                    // Replace [i] with current loop index
                    // Only require currentIndex >= 0, don't require array name match for table loops
                    // (Table loops are simpler - one array per row)
                    if (currentIndex >= 0) {
                        // For table loops, we're always in the correct context
                        // Replace [i] with the actual numeric index
                        targetIndex = currentIndex;
                        System.out.println("DEBUG: Setting targetIndex to " + targetIndex);
                    } else {
                        // Not in a loop context (currentIndex < 0), cannot replace [i]
                        // Still add to replacements with empty string so it gets replaced
                        System.out.println("WARN: currentIndex < 0 for placeholder with [i]: " + fullPlaceholder + " - Replacing with empty string");
                        allReplacements.put(fullPlaceholder, "");
                        continue;
                    }
                } else {
                    // Use the numeric index directly (e.g., [0], [1], [2])
                    try {
                        targetIndex = Integer.parseInt(indexStr);
                        System.out.println("DEBUG: Parsed numeric index: " + targetIndex);
                    } catch (NumberFormatException e) {
                        // Invalid index format, still replace with empty string
                        System.out.println("WARN: Invalid index format: " + indexStr + " for placeholder: " + fullPlaceholder + " - Replacing with empty string");
                        allReplacements.put(fullPlaceholder, "");
                        continue;
                    }
                }
                
                // Validate index bounds if itemsList is provided
                if (itemsList != null) {
                    if (targetIndex < 0 || targetIndex >= itemsList.size()) {
                        System.out.println("WARN: Index out of bounds: " + targetIndex + " for array size: " + itemsList.size() + " for placeholder: " + fullPlaceholder + " - Replacing with empty string");
                        allReplacements.put(fullPlaceholder, "");
                        continue;
                    }
                }
                
                // Use JSON-Path to extract value: $.arrayName[index].property
                // Replace [i] with actual numeric index
                String jsonPathExpr = "$." + arrayName + "[" + targetIndex + "]." + property;
                System.out.println("DEBUG: Replacing [i] with " + targetIndex + ", JSON-Path: " + jsonPathExpr + ", Placeholder: " + fullPlaceholder);
                
                try {
                    Object value = originalJsonContext.read(jsonPathExpr);
                    String valueStr = value != null ? value.toString() : "";
                    allReplacements.put(fullPlaceholder, valueStr);
                    System.out.println("DEBUG: Extracted value: " + valueStr + " for placeholder: " + fullPlaceholder);
                } catch (PathNotFoundException e) {
                    // Path not found - replace with empty string and log
                    System.out.println("WARN: JSON-Path not found: " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - Replacing with empty string");
                    allReplacements.put(fullPlaceholder, "");
                } catch (Exception e) {
                    // Any other error - replace with empty string and log
                    System.out.println("ERROR: Error reading JSON-Path " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - " + e.getMessage() + " - Replacing with empty string");
                    allReplacements.put(fullPlaceholder, "");
                }
            }
        }

        // Step 2: Process radio buttons using JSON-Path
        if (hasRadios) {
            radioMatcher.reset();
            while (radioMatcher.find()) {
                String radioGroupPath = radioMatcher.group(1);
                String radioValue = radioMatcher.group(2);
                String fullPlaceholder = radioMatcher.group(0);
                
                String radioGroup = extractVariableName(radioGroupPath);
                String jsonPathExpr = "$." + radioGroup;
                try {
                    Object groupValueObj = jsonContext.read(jsonPathExpr);
                    String groupValue = groupValueObj != null ? groupValueObj.toString() : null;
                    boolean isSelected = radioValue.equalsIgnoreCase(groupValue);
                    String radioSymbol = isSelected ? "◉" : "○";
                    allReplacements.put(fullPlaceholder, radioSymbol);
                } catch (PathNotFoundException e) {
                    allReplacements.put(fullPlaceholder, "○");
                }
            }
        }

        // Step 3: Process checkboxes using JSON-Path
        if (hasCheckboxes) {
            checkboxMatcher.reset();
            while (checkboxMatcher.find()) {
                String checkboxGroupPath = checkboxMatcher.group(1);
                String checkboxValue = checkboxMatcher.group(2);
                String fullPlaceholder = checkboxMatcher.group(0);
                
                String checkboxGroup = extractVariableName(checkboxGroupPath);
                String jsonPathExpr = "$." + checkboxGroup;
                try {
                    Object groupValueObj = jsonContext.read(jsonPathExpr);
                    String groupValue = groupValueObj != null ? groupValueObj.toString() : null;
                    boolean isChecked = checkboxValue.equalsIgnoreCase(groupValue);
                    String checkboxSymbol = isChecked ? "☒" : "☐";
                    allReplacements.put(fullPlaceholder, checkboxSymbol);
                } catch (PathNotFoundException e) {
                    allReplacements.put(fullPlaceholder, "☐");
                }
            }
        }

        // Step 4: Process regular placeholders using JSON-Path
        if (hasRegularPlaceholders) {
            regularMatcher.reset();
            while (regularMatcher.find()) {
                String variablePath = regularMatcher.group(1);
                String fullPlaceholder = regularMatcher.group(0);

                // Skip if this placeholder contains [i] or [number] - it's a loop item placeholder
                if (fullPlaceholder.matches(".*\\[([i]|\\d+)\\].*")) {
                    continue;
                }
                
                if (allReplacements.containsKey(fullPlaceholder)) {
                    continue;
                }

                // Convert variable path to JSON-Path expression
                // e.g., "$orderId" -> "$.orderId", "$customer.name" -> "$.customer.name"
                String variableName = extractVariableName(variablePath);
                String jsonPathExpr = convertToJsonPath(variableName);
                
                try {
                    Object valueObj = jsonContext.read(jsonPathExpr);
                    String value = valueObj != null ? valueObj.toString() : "";
                    allReplacements.put(fullPlaceholder, value);
                } catch (PathNotFoundException e) {
                    // Path not found - replace with empty string and log
                    System.out.println("WARN: JSON-Path not found: " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - Replacing with empty string");
                    allReplacements.put(fullPlaceholder, "");
                } catch (Exception e) {
                    // Any other error - replace with empty string and log
                    System.out.println("ERROR: Error reading JSON-Path " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - " + e.getMessage() + " - Replacing with empty string");
                    allReplacements.put(fullPlaceholder, "");
                }
            }
        }

        if (!allReplacements.isEmpty()) {
            processRunsWithPlaceholderFormatting(paragraph, allReplacements, currentIndex);
        }
    }

    /**
     * Converts a variable path to JSON-Path expression
     * Examples:
     *   "orderId" -> "$.orderId"
     *   "customer.name" -> "$.customer.name"
     *   "items[0].name" -> "$.items[0].name"
     */
    private String convertToJsonPath(String variablePath) {
        if (variablePath == null || variablePath.isEmpty()) {
            return "$";
        }
        
        // If it already starts with $, return as is
        if (variablePath.startsWith("$")) {
            return variablePath;
        }
        
        // Otherwise, prepend $.
        return "$." + variablePath;
    }

    // ==================== Table Loop Implementation Methods ====================

    /**
     * Processes table row loops using JSON-Path
     * Supports two syntaxes:
     * 1. Single row: {{#loop:$arrayName}} and {{#/loop}} in the same row
     * 2. Multiple rows: {{#loop:$arrayName}} in one row, {{#/loop}} in a different row
     * 
     * The row(s) will be duplicated for each item in the array
     * Use {{$arrayName[i].property}} in cells to access item properties
     * 
     * Examples:
     * Single row: | {{#loop:$orderItems}} | {{$orderItems[i].itemName}} | {{$orderItems[i].quantity}} | {{#/loop}} |
     * Multiple rows: 
     *   Row 1: | {{#loop:$orderItems}} | Header 1 | Header 2 |
     *   Row 2: | {{$orderItems[i].itemName}} | {{$orderItems[i].quantity}} | {{$orderItems[i].unitPrice}} |
     *   Row 3: | Footer | {{#/loop}} | Total |
     */
    private void processTableLoopsWithJsonPath(XWPFTable table, DocumentContext jsonContext) {
        // Get a snapshot of rows to avoid concurrent modification issues
        List<XWPFTableRow> rows = new ArrayList<>(table.getRows());
        
        // First, find and process multi-row loops (from end to start to avoid index shifting)
        List<MultiRowLoopInfo> multiRowLoops = findMultiRowLoops(rows);
        for (MultiRowLoopInfo multiRowLoop : multiRowLoops) {
            processMultiRowLoop(table, multiRowLoop, jsonContext);
        }
        
        // Then, process single-row loops (from end to start to avoid index shifting)
        // Refresh rows list as multi-row loops may have modified the table
        rows = new ArrayList<>(table.getRows());
        for (int rowIndex = rows.size() - 1; rowIndex >= 0; rowIndex--) {
            try {
                XWPFTableRow row = rows.get(rowIndex);
                if (row == null) {
                    continue;
                }
                
                // Check if this row contains both loop start and end markers (single-row loop)
                LoopMarkerInfo markerInfo = findLoopMarkersInRow(row);
                if (markerInfo != null) {
                    // Process this single-row loop
                    processTableLoopRow(table, row, rowIndex, markerInfo, jsonContext);
                }
            } catch (IndexOutOfBoundsException e) {
                // Row index out of bounds, skip this row
                continue;
            }
        }
    }

    /**
     * Information about loop markers in a table row
     */
    private static class LoopMarkerInfo {
        String arrayName;
        int startCellIndex;  // Cell index containing {{#loop:$arrayName}}
        int endCellIndex;    // Cell index containing {{#/loop}}
        
        LoopMarkerInfo(String arrayName, int startCellIndex, int endCellIndex) {
            this.arrayName = arrayName;
            this.startCellIndex = startCellIndex;
            this.endCellIndex = endCellIndex;
        }
    }
    
    /**
     * Information about loop markers spanning multiple rows
     */
    private static class MultiRowLoopInfo {
        String arrayName;
        int startRowIndex;   // Row index containing {{#loop:$arrayName}}
        int endRowIndex;     // Row index containing {{#/loop}}
        int startCellIndex;  // Cell index in start row containing {{#loop:$arrayName}}
        int endCellIndex;    // Cell index in end row containing {{#/loop}}
        
        MultiRowLoopInfo(String arrayName, int startRowIndex, int endRowIndex, int startCellIndex, int endCellIndex) {
            this.arrayName = arrayName;
            this.startRowIndex = startRowIndex;
            this.endRowIndex = endRowIndex;
            this.startCellIndex = startCellIndex;
            this.endCellIndex = endCellIndex;
        }
    }

    /**
     * Finds loop markers {{#loop:$arrayName}} and {{#/loop}} in a table row
     * Returns LoopMarkerInfo if both markers are found, null otherwise
     */
    private LoopMarkerInfo findLoopMarkersInRow(XWPFTableRow row) {
        if (row == null) {
            return null;
        }
        
        String arrayName = null;
        int startCellIndex = -1;
        int endCellIndex = -1;
        
        List<XWPFTableCell> cells = row.getTableCells();
        if (cells == null || cells.isEmpty()) {
            return null;
        }
        
        for (int cellIndex = 0; cellIndex < cells.size(); cellIndex++) {
            try {
                XWPFTableCell cell = cells.get(cellIndex);
                if (cell == null) {
                    continue;
                }
                
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    if (paragraph == null) {
                        continue;
                    }
                    
                    String text = getParagraphText(paragraph);
                    if (text != null) {
                        // Check for loop start marker
                        Matcher startMatcher = LOOP_START_PATTERN.matcher(text);
                        if (startMatcher.find()) {
                            String loopVariablePath = startMatcher.group(1);
                            arrayName = extractVariableName(loopVariablePath);
                            startCellIndex = cellIndex;
                        }
                        
                        // Check for loop end marker
                        Matcher endMatcher = LOOP_END_PATTERN.matcher(text);
                        if (endMatcher.find()) {
                            endCellIndex = cellIndex;
                        }
                    }
                }
            } catch (IndexOutOfBoundsException e) {
                // Skip this cell if index is out of bounds
                continue;
            }
        }
        
        // Both markers must be found and indices must be valid
        if (arrayName != null && startCellIndex >= 0 && endCellIndex >= 0 
            && startCellIndex < cells.size() && endCellIndex < cells.size()) {
            return new LoopMarkerInfo(arrayName, startCellIndex, endCellIndex);
        }
        
        return null;
    }

    /**
     * Finds loop markers that span multiple rows
     * Returns a list of MultiRowLoopInfo sorted by start row index (descending for safe processing)
     */
    private List<MultiRowLoopInfo> findMultiRowLoops(List<XWPFTableRow> rows) {
        List<MultiRowLoopInfo> multiRowLoops = new ArrayList<>();
        
        // Find all start markers and end markers
        Map<Integer, String> startMarkers = new HashMap<>(); // rowIndex -> arrayName
        Map<Integer, Integer> startCellIndices = new HashMap<>(); // rowIndex -> cellIndex
        Map<Integer, Integer> endCellIndices = new HashMap<>(); // rowIndex -> cellIndex
        
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = rows.get(rowIndex);
            if (row == null) {
                continue;
            }
            
            List<XWPFTableCell> cells = row.getTableCells();
            if (cells == null || cells.isEmpty()) {
                continue;
            }
            
            for (int cellIndex = 0; cellIndex < cells.size(); cellIndex++) {
                try {
                    XWPFTableCell cell = cells.get(cellIndex);
                    if (cell == null) {
                        continue;
                    }
                    
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        if (paragraph == null) {
                            continue;
                        }
                        
                        String text = getParagraphText(paragraph);
                        if (text != null) {
                            // Check for loop start marker
                            Matcher startMatcher = LOOP_START_PATTERN.matcher(text);
                            if (startMatcher.find()) {
                                String loopVariablePath = startMatcher.group(1);
                                String arrayName = extractVariableName(loopVariablePath);
                                startMarkers.put(rowIndex, arrayName);
                                startCellIndices.put(rowIndex, cellIndex);
                            }
                            
                            // Check for loop end marker
                            Matcher endMatcher = LOOP_END_PATTERN.matcher(text);
                            if (endMatcher.find()) {
                                endCellIndices.put(rowIndex, cellIndex);
                            }
                        }
                    }
                } catch (Exception e) {
                    continue;
                }
            }
        }
        
        // Match start markers with end markers (find the closest end marker after each start marker)
        for (Map.Entry<Integer, String> startEntry : startMarkers.entrySet()) {
            int startRowIndex = startEntry.getKey();
            String arrayName = startEntry.getValue();
            int startCellIndex = startCellIndices.get(startRowIndex);
            
            // Find the closest end marker after this start marker
            int endRowIndex = -1;
            int endCellIndex = -1;
            
            for (Map.Entry<Integer, Integer> endEntry : endCellIndices.entrySet()) {
                int candidateEndRow = endEntry.getKey();
                if (candidateEndRow > startRowIndex) {
                    // Check if this row doesn't have a start marker (to avoid matching with nested loops)
                    // For simplicity, we'll match with the first end marker after start
                    if (endRowIndex == -1 || candidateEndRow < endRowIndex) {
                        endRowIndex = candidateEndRow;
                        endCellIndex = endEntry.getValue();
                    }
                }
            }
            
            // Only create multi-row loop if end marker is in a different row
            if (endRowIndex > startRowIndex) {
                // Check if this is not a single-row loop (both markers in same row)
                LoopMarkerInfo singleRowCheck = findLoopMarkersInRow(rows.get(startRowIndex));
                if (singleRowCheck == null) {
                    // This is a multi-row loop
                    multiRowLoops.add(new MultiRowLoopInfo(arrayName, startRowIndex, endRowIndex, startCellIndex, endCellIndex));
                }
            }
        }
        
        // Sort by start row index descending for safe processing (process from end to start)
        multiRowLoops.sort((a, b) -> Integer.compare(b.startRowIndex, a.startRowIndex));
        
        return multiRowLoops;
    }

    /**
     * Processes a multi-row loop
     * Clones all rows between start and end markers for each item in the array
     */
    @SuppressWarnings("unchecked")
    private void processMultiRowLoop(XWPFTable table, MultiRowLoopInfo loopInfo, DocumentContext jsonContext) {
        // Get the array from JSON using JSON-Path
        String jsonPathExpr = "$." + loopInfo.arrayName;
        Object listObj;
        try {
            listObj = jsonContext.read(jsonPathExpr);
        } catch (PathNotFoundException e) {
            // Array not found, remove the loop markers and return
            removeLoopMarkersFromMultiRow(table, loopInfo);
            return;
        }
        
        if (!(listObj instanceof List)) {
            // Not a list, remove the loop markers and return
            removeLoopMarkersFromMultiRow(table, loopInfo);
            return;
        }
        
        List<Object> items = (List<Object>) listObj;
        if (items.isEmpty()) {
            // Empty list, remove all rows between start and end (inclusive)
            List<XWPFTableRow> rows = table.getRows();
            for (int i = loopInfo.endRowIndex; i >= loopInfo.startRowIndex; i--) {
                if (i >= 0 && i < rows.size()) {
                    try {
                        table.removeRow(i);
                    } catch (Exception e) {
                        // Skip if removal fails
                    }
                }
            }
            return;
        }
        
        // Get all rows between start and end (inclusive)
        List<XWPFTableRow> templateRows = new ArrayList<>();
        List<String> templateRowXmls = new ArrayList<>();
        List<XWPFTableRow> rows = table.getRows();
        
        for (int i = loopInfo.startRowIndex; i <= loopInfo.endRowIndex && i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            if (row != null) {
                templateRows.add(row);
                templateRowXmls.add(row.getCtRow().xmlText());
            }
        }
        
        if (templateRows.isEmpty()) {
            return;
        }
        
        // Process the first set of rows (index 0) - remove markers and process placeholders
        for (int rowIdx = 0; rowIdx < templateRows.size(); rowIdx++) {
            XWPFTableRow row = templateRows.get(rowIdx);
            LoopMarkerInfo markerInfo = new LoopMarkerInfo(loopInfo.arrayName, 
                (rowIdx == 0) ? loopInfo.startCellIndex : -1,
                (rowIdx == templateRows.size() - 1) ? loopInfo.endCellIndex : -1);
            processTableRowForItem(row, jsonContext, loopInfo.arrayName, 0, jsonContext, items, markerInfo);
        }
        
        // Clone and insert additional sets of rows for remaining items
        if (items.size() > 1) {
            int insertPosition = loopInfo.endRowIndex + 1;
            
            for (int itemIndex = 1; itemIndex < items.size(); itemIndex++) {
                // Clone all template rows for this item
                for (int rowIdx = 0; rowIdx < templateRowXmls.size(); rowIdx++) {
                    try {
                        String templateRowXml = templateRowXmls.get(rowIdx);
                        
                        // Parse the cloned row XML
                        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow clonedRowCT =
                            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow.Factory.parse(templateRowXml);
                        
                        // Get current table size before insertion
                        int currentTableSize = table.getCTTbl().sizeOfTrArray();
                        
                        // Validate insert position
                        if (insertPosition > currentTableSize) {
                            insertPosition = currentTableSize;
                        }
                        if (insertPosition < 0) {
                            insertPosition = 0;
                        }
                        
                        // Insert a new empty row at the specified position
                        table.getCTTbl().insertNewTr(insertPosition);
                        
                        // Replace the empty row with our cloned row content
                        table.getCTTbl().setTrArray(insertPosition, clonedRowCT);
                        
                        // Get the CTRow we just set
                        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow rowCT = 
                            table.getCTTbl().getTrArray(insertPosition);
                        if (rowCT == null) {
                            insertPosition++;
                            continue;
                        }
                        
                        // Create XWPFTableRow wrapper from the CTRow
                        XWPFTableRow newRow = new XWPFTableRow(rowCT, table);
                        
                        // Process the new row for this item
                        LoopMarkerInfo markerInfo = new LoopMarkerInfo(loopInfo.arrayName,
                            (rowIdx == 0) ? loopInfo.startCellIndex : -1,
                            (rowIdx == templateRowXmls.size() - 1) ? loopInfo.endCellIndex : -1);
                        processTableRowForItem(newRow, jsonContext, loopInfo.arrayName, itemIndex, jsonContext, items, markerInfo);
                        
                        insertPosition++;
                    } catch (Exception e) {
                        System.out.println("DEBUG: Exception cloning multi-row loop row: " + e.getMessage());
                        e.printStackTrace();
                        continue;
                    }
                }
            }
        }
    }
    
    /**
     * Removes loop markers from a multi-row loop
     */
    private void removeLoopMarkersFromMultiRow(XWPFTable table, MultiRowLoopInfo loopInfo) {
        List<XWPFTableRow> rows = table.getRows();
        if (loopInfo.startRowIndex >= 0 && loopInfo.startRowIndex < rows.size()) {
            XWPFTableRow startRow = rows.get(loopInfo.startRowIndex);
            if (startRow != null) {
                LoopMarkerInfo startMarkerInfo = new LoopMarkerInfo(loopInfo.arrayName, loopInfo.startCellIndex, -1);
                removeLoopMarkersFromRow(startRow, startMarkerInfo);
            }
        }
        if (loopInfo.endRowIndex >= 0 && loopInfo.endRowIndex < rows.size()) {
            XWPFTableRow endRow = rows.get(loopInfo.endRowIndex);
            if (endRow != null) {
                LoopMarkerInfo endMarkerInfo = new LoopMarkerInfo(loopInfo.arrayName, -1, loopInfo.endCellIndex);
                removeLoopMarkersFromRow(endRow, endMarkerInfo);
            }
        }
    }

    /**
     * Processes a single table loop row
     * Duplicates the row for each item in the array and replaces placeholders
     * Removes both loop markers after processing
     */
    @SuppressWarnings("unchecked")
    private void processTableLoopRow(XWPFTable table, XWPFTableRow templateRow, int rowIndex, 
                                     LoopMarkerInfo markerInfo, DocumentContext jsonContext) {
        // Get the array from JSON using JSON-Path
        String jsonPathExpr = "$." + markerInfo.arrayName;
        Object listObj;
        try {
            listObj = jsonContext.read(jsonPathExpr);
        } catch (PathNotFoundException e) {
            // Array not found, remove the loop markers and return
            removeLoopMarkersFromRow(templateRow, markerInfo);
            return;
        }
        
        if (!(listObj instanceof List)) {
            // Not a list, remove the loop markers and return
            removeLoopMarkersFromRow(templateRow, markerInfo);
            return;
        }
        
        List<Object> items = (List<Object>) listObj;
        if (items.isEmpty()) {
            // Empty list, remove the template row entirely
            // Verify the row index is still valid before removing
            List<XWPFTableRow> currentRows = table.getRows();
            if (rowIndex >= 0 && rowIndex < currentRows.size()) {
                try {
                    table.removeRow(rowIndex);
                } catch (IndexOutOfBoundsException e) {
                    // Row index invalid, skip removal
                }
            }
            return;
        }
        
        // Step 1: Clone the template row XML BEFORE any modifications (preserve original structure)
        String templateRowXml = templateRow.getCtRow().xmlText();
        
        // Step 2: Process the template row for the first item (index 0)
        // This will: 1) Remove loop markers, 2) Replace [i] with 0, 3) Extract values
        processTableRowForItem(templateRow, jsonContext, markerInfo.arrayName, 0, jsonContext, items, markerInfo);
        
        // Step 3: Insert additional rows for remaining items (if any)
        if (items.size() > 1) {
            System.out.println("DEBUG: Starting to clone rows for " + (items.size() - 1) + " additional items");
            // Start inserting after the template row (which is now at rowIndex after processing)
            int insertPosition = rowIndex + 1;
            
            for (int itemIndex = 1; itemIndex < items.size(); itemIndex++) {
                System.out.println("DEBUG: Processing item index " + itemIndex + ", insertPosition=" + insertPosition);
                try {
                    // Clone the original template row XML (with markers and placeholders)
                    org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow clonedRowCT =
                        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow.Factory.parse(templateRowXml);
                    
                    System.out.println("DEBUG: Successfully parsed cloned row XML");
                    
                    // Get current table size before insertion
                    int currentTableSize = table.getCTTbl().sizeOfTrArray();
                    System.out.println("DEBUG: Current table size: " + currentTableSize);
                    
                    // Validate insert position
                    if (insertPosition > currentTableSize) {
                        System.out.println("DEBUG: Adjusting insertPosition from " + insertPosition + " to " + currentTableSize);
                        insertPosition = currentTableSize;
                    }
                    if (insertPosition < 0) {
                        System.out.println("DEBUG: Adjusting insertPosition from " + insertPosition + " to 0");
                        insertPosition = 0;
                    }
                    
                    System.out.println("DEBUG: Final insertPosition: " + insertPosition);
                    
                    // Insert a new empty row at the specified position
                    table.getCTTbl().insertNewTr(insertPosition);
                    System.out.println("DEBUG: Inserted new empty row at position " + insertPosition);
                    
                    // Replace the empty row with our cloned row content
                    table.getCTTbl().setTrArray(insertPosition, clonedRowCT);
                    System.out.println("DEBUG: Set cloned row content at position " + insertPosition);
                    
                    // Get the CTRow we just set (similar to how paragraphs are handled)
                    org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow rowCT = table.getCTTbl().getTrArray(insertPosition);
                    if (rowCT == null) {
                        System.out.println("DEBUG: ERROR - Could not get CTRow at position " + insertPosition);
                        insertPosition++;
                        continue;
                    }
                    
                    System.out.println("DEBUG: Retrieved CTRow at position " + insertPosition);
                    
                    // Create XWPFTableRow wrapper from the CTRow (similar to XWPFParagraph wrapper)
                    XWPFTableRow newRow = new XWPFTableRow(rowCT, table);
                    System.out.println("DEBUG: Created XWPFTableRow wrapper");
                    
                    // Process the new row for this item: remove markers, replace [i], extract values
                    System.out.println("DEBUG: Calling processTableRowForItem for itemIndex=" + itemIndex);
                    processTableRowForItem(newRow, jsonContext, markerInfo.arrayName, itemIndex, jsonContext, items, markerInfo);
                    System.out.println("DEBUG: Completed processTableRowForItem for itemIndex=" + itemIndex);
                    
                    // Move to next position for next iteration
                    insertPosition++;
                    System.out.println("DEBUG: Incremented insertPosition to " + insertPosition);
                    
                } catch (org.apache.xmlbeans.XmlException e) {
                    // If cloning/parsing fails, skip this item
                    System.out.println("DEBUG: XmlException while cloning row: " + e.getMessage());
                    e.printStackTrace();
                    continue;
                } catch (IndexOutOfBoundsException e) {
                    // If row access fails, skip this item
                    System.out.println("DEBUG: IndexOutOfBoundsException while accessing row: " + e.getMessage());
                    e.printStackTrace();
                    continue;
                } catch (Exception e) {
                    // Catch any other exceptions
                    System.out.println("DEBUG: Exception while processing row clone: " + e.getMessage());
                    e.printStackTrace();
                    continue;
                }
            }
            System.out.println("DEBUG: Finished cloning all additional rows");
        }
    }

    /**
     * Processes a table row for a specific item in the loop
     * 1. Removes loop markers ({{#loop:$arrayName}} and {{#/loop}})
     * 2. Replaces [i] with the actual numeric index
     * 3. Extracts and replaces all placeholders with values from JSON
     */
    private void processTableRowForItem(XWPFTableRow row, DocumentContext jsonContext,
                                        String loopArrayName, int currentIndex,
                                        DocumentContext originalJsonContext, 
                                        List<Object> itemsList,
                                        LoopMarkerInfo markerInfo) {
        if (row == null) {
            return;
        }
        
        List<XWPFTableCell> cells = row.getTableCells();
        if (cells == null || cells.isEmpty()) {
            return;
        }
        
        // Debug: Log that we're processing this row
        System.out.println("DEBUG: Processing table row for item index " + currentIndex + ", array: " + loopArrayName);
        
        // Process each cell: remove markers AND process placeholders in one pass
        // This ensures placeholders are processed immediately after marker removal
        for (XWPFTableCell cell : cells) {
            if (cell == null) {
                continue;
            }
            
            try {
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                if (paragraphs == null || paragraphs.isEmpty()) {
                    // If no paragraphs, try to get text directly from cell
                    String cellText = cell.getText();
                    if (cellText != null && (LOOP_START_PATTERN.matcher(cellText).find() || LOOP_END_PATTERN.matcher(cellText).find())) {
                        // Create a paragraph to process
                        XWPFParagraph para = cell.addParagraph();
                        XWPFRun run = para.createRun();
                        run.setText(cellText);
                        // Remove markers first
                        removePlaceholderFromParagraph(para, LOOP_START_PATTERN);
                        removePlaceholderFromParagraph(para, LOOP_END_PATTERN);
                        // Then process placeholders
                        processAllPlaceholdersWithJsonPath(para, jsonContext, loopArrayName, currentIndex, 
                                                          originalJsonContext, itemsList);
                    }
                    continue;
                }
                
                for (XWPFParagraph paragraph : paragraphs) {
                    if (paragraph == null) {
                        continue;
                    }
                    
                    String textBefore = getParagraphText(paragraph);
                    if (textBefore == null || textBefore.trim().isEmpty()) {
                        continue;
                    }
                    
                    // Debug: Log text before processing
                    System.out.println("DEBUG: Cell text before processing: " + textBefore);
                    
                    // Step 1: Remove loop markers (but preserve other text including placeholders)
                    removePlaceholderFromParagraph(paragraph, LOOP_START_PATTERN);
                    removePlaceholderFromParagraph(paragraph, LOOP_END_PATTERN);
                    
                    // Double-check: if markers still exist, remove them more aggressively
                    String textAfterMarkers = getParagraphText(paragraph);
                    if (textAfterMarkers != null) {
                        boolean hasStartMarker = LOOP_START_PATTERN.matcher(textAfterMarkers).find();
                        boolean hasEndMarker = LOOP_END_PATTERN.matcher(textAfterMarkers).find();
                        
                        if (hasStartMarker || hasEndMarker) {
                            // More aggressive removal - replace all markers directly
                            String cleaned = textAfterMarkers;
                            cleaned = LOOP_START_PATTERN.matcher(cleaned).replaceAll("");
                            cleaned = LOOP_END_PATTERN.matcher(cleaned).replaceAll("");
                            cleaned = cleaned.trim();
                            
                            if (!cleaned.equals(textAfterMarkers)) {
                                // Clear and recreate
                                int runsCount = paragraph.getRuns().size();
                                for (int i = runsCount - 1; i >= 0; i--) {
                                    paragraph.removeRun(i);
                                }
                                if (!cleaned.isEmpty()) {
                                    paragraph.createRun().setText(cleaned);
                                }
                                System.out.println("DEBUG: Aggressively removed markers, new text: " + cleaned);
                            }
                        }
                    }
                    
                    // Step 2: Process placeholders (replaces [i] with currentIndex and extracts values)
                    // IMPORTANT: This must happen AFTER marker removal to ensure placeholders are still in the text
                    String textBeforePlaceholders = getParagraphText(paragraph);
                    System.out.println("DEBUG: Text before placeholder processing: " + textBeforePlaceholders);
                    
                    if (textBeforePlaceholders != null && !textBeforePlaceholders.trim().isEmpty()) {
                        // Process all placeholders in this paragraph
                        // IMPORTANT: Pass loopArrayName, currentIndex, and itemsList
                        // This allows processAllPlaceholdersWithJsonPath to replace [i] with currentIndex
                        processAllPlaceholdersWithJsonPath(paragraph, jsonContext, loopArrayName, currentIndex, 
                                                          originalJsonContext, itemsList);
                        
                        // Debug: Log paragraph text after processing
                        String textAfterPlaceholders = getParagraphText(paragraph);
                        System.out.println("DEBUG: Text after placeholder processing: " + textAfterPlaceholders);
                    }
                }
            } catch (Exception e) {
                System.out.println("DEBUG: Exception processing cell: " + e.getMessage());
                e.printStackTrace();
                continue;
            }
        }
    }

    /**
     * Removes both loop markers {{#loop:$arrayName}} and {{#/loop}} from a table row
     * This ensures the loop markers don't appear in the final document
     * Searches ALL cells to find and remove markers (not just start/end cells)
     */
    private void removeLoopMarkersFromRow(XWPFTableRow row, LoopMarkerInfo markerInfo) {
        if (row == null) {
            return;
        }
        
        List<XWPFTableCell> cells = row.getTableCells();
        if (cells == null || cells.isEmpty()) {
            return;
        }
        
        // Search ALL cells for loop markers and remove them
        // This ensures markers are removed even if cell structure changes after cloning
        for (XWPFTableCell cell : cells) {
            if (cell == null) {
                continue;
            }
            
            try {
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                if (paragraphs == null || paragraphs.isEmpty()) {
                    continue;
                }
                
                for (XWPFParagraph paragraph : paragraphs) {
                    if (paragraph != null) {
                        // Remove loop start marker if found
                        removePlaceholderFromParagraph(paragraph, LOOP_START_PATTERN);
                        // Remove loop end marker if found
                        removePlaceholderFromParagraph(paragraph, LOOP_END_PATTERN);
                    }
                }
            } catch (Exception e) {
                // If processing a cell fails, continue with next cell
                continue;
            }
        }
    }

    /**
     * Processes placeholders in a table row using JSON-Path
     * Handles loop item placeholders like {{$arrayName[i].property}}
     * Replaces [i] with the actual numeric index (0, 1, 2, etc.)
     */
    private void processTableRowPlaceholdersWithJsonPath(XWPFTableRow row, DocumentContext jsonContext,
                                                          String loopArrayName, int currentIndex,
                                                          DocumentContext originalJsonContext, 
                                                          List<Object> itemsList) {
        if (row == null) {
            return;
        }
        
        try {
            List<XWPFTableCell> cells = row.getTableCells();
            if (cells == null || cells.isEmpty()) {
                return;
            }
            
            // Process each cell in the row
            for (XWPFTableCell cell : cells) {
                if (cell == null) {
                    continue;
                }
                
                try {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    if (paragraphs == null || paragraphs.isEmpty()) {
                        // If no paragraphs, the cell might be empty - continue to next cell
                        continue;
                    }
                    
                    // Process each paragraph in the cell
                for (XWPFParagraph paragraph : paragraphs) {
                    if (paragraph != null) {
                        String paraText = getParagraphText(paragraph);
                        if (paraText != null && !paraText.trim().isEmpty()) {
                            // Debug: Log paragraph text before processing
                            System.out.println("DEBUG: Processing paragraph text: " + paraText);
                            
                            // Process all placeholders in this paragraph
                            // IMPORTANT: Pass loopArrayName, currentIndex, and itemsList
                            // This allows processAllPlaceholdersWithJsonPath to replace [i] with currentIndex
                            processAllPlaceholdersWithJsonPath(paragraph, jsonContext, loopArrayName, currentIndex, 
                                                              originalJsonContext, itemsList);
                            
                            // Debug: Log paragraph text after processing
                            String paraTextAfter = getParagraphText(paragraph);
                            System.out.println("DEBUG: After processing: " + paraTextAfter);
                        }
                    }
                }
                } catch (Exception e) {
                    // If processing a cell fails, continue with next cell
                    // Don't let one cell failure stop the entire row processing
                    continue;
                }
            }
        } catch (Exception e) {
            // If processing the row fails, log and continue
            // Don't throw - just skip this row to avoid breaking the entire document processing
            System.err.println("Error in processTableRowPlaceholdersWithJsonPath: " + e.getMessage());
        }
    }
}


