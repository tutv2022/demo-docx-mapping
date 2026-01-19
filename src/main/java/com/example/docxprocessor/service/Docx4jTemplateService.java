package com.example.docxprocessor.service;

import com.jayway.jsonpath.DocumentContext;
import com.jayway.jsonpath.JsonPath;
import com.jayway.jsonpath.PathNotFoundException;
import jakarta.xml.bind.JAXBElement;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.XmlUtils;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.namespace.QName;

/**
 * docx4j-based implementation of the DOCX template processor.
 *
 * Supported syntax:
 * - {{$path}} (JSONPath: "$.path")
 * - {{$array[i].prop}} loop-item placeholder (requires loop context)
 * - {{#loop:$array}} ... {{#/loop}} in tables (can span multiple rows)
 * - {{#checkbox:$var:value}} and {{#radio:$var:value}}
 *
 * Notes:
 * - Table loops and body paragraph loops are supported.
 * - Placeholders split across runs are handled by operating on concatenated w:t text per paragraph.
 */
@Service
public class Docx4jTemplateService {

    private static final String WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final QName QNAME_TR = new QName(WML_NS, "tr");

    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\{\\{\\$([^}]+)\\}\\}");
    private static final Pattern CHECKBOX_PATTERN = Pattern.compile("\\{\\{#checkbox:\\s*\\$([^:]+):([^}]+)\\}\\}");
    private static final Pattern RADIO_PATTERN = Pattern.compile("\\{\\{#radio:\\s*\\$([^:]+):([^}]+)\\}\\}");
    // Allow whitespace around markers to tolerate Word run splits/spaces
    private static final Pattern LOOP_START_PATTERN = Pattern.compile("\\{\\{\\s*#loop:\\s*\\$([^}]+)\\s*\\}\\}");
    private static final Pattern LOOP_END_PATTERN = Pattern.compile("\\{\\{\\s*#/loop\\s*\\}\\}");
    private static final Pattern LOOP_ITEM_PATTERN = Pattern.compile("\\{\\{\\s*\\$([^\\[\\}]+)\\[([^\\]]+)\\]\\.([^\\}]+)\\s*\\}\\}");

    public byte[] processTemplatePreservingFormat(InputStream templateInputStream, Map<String, Object> data) throws IOException {
        WordprocessingMLPackage pkg = loadPackage(templateInputStream);
        forEachParagraph(pkg, p -> processParagraphWithMap(p, data));
        return savePackage(pkg);
    }

    public byte[] processTemplateWithJsonPath(InputStream templateInputStream, String jsonString) throws IOException {
        return processTemplateWithJsonPath(templateInputStream, JsonPath.parse(jsonString));
    }

    public byte[] processTemplateWithJsonPath(InputStream templateInputStream, DocumentContext jsonContext) throws IOException {
        WordprocessingMLPackage pkg = loadPackage(templateInputStream);
        processParagraphLoopsWithJsonPath(pkg, jsonContext);
        forEachTable(pkg, tbl -> processTableLoopsWithJsonPath(tbl, jsonContext));
        forEachParagraph(pkg, p -> processAllPlaceholdersWithJsonPath(p, jsonContext, null, -1, jsonContext, null));
        return savePackage(pkg);
    }

    /**
     * Generate a PDF by applying JSON-Path placeholder replacement, then exporting DOCX->PDF via XSL-FO.
     */
    public byte[] processTemplateWithJsonPathToPdf(InputStream templateInputStream, String jsonString) throws IOException {
        return processTemplateWithJsonPathToPdf(templateInputStream, JsonPath.parse(jsonString));
    }

    /**
     * Generate a PDF by applying JSON-Path placeholder replacement, then exporting DOCX->PDF via XSL-FO.
     */
    public byte[] processTemplateWithJsonPathToPdf(InputStream templateInputStream, DocumentContext jsonContext) throws IOException {
        WordprocessingMLPackage pkg = loadPackage(templateInputStream);
        processParagraphLoopsWithJsonPath(pkg, jsonContext);
        forEachTable(pkg, tbl -> processTableLoopsWithJsonPath(tbl, jsonContext));
        forEachParagraph(pkg, p -> processAllPlaceholdersWithJsonPath(p, jsonContext, null, -1, jsonContext, null));

        try (ByteArrayOutputStream pdfOut = new ByteArrayOutputStream()) {
            FOSettings foSettings = Docx4J.createFOSettings();
            // Use setOpcPackage (preferred) so docx4j initializes FOP config correctly
            foSettings.setOpcPackage(pkg);
            foSettings.setApacheFopMime(FOSettings.MIME_PDF);
            Docx4J.toFO(foSettings, pdfOut, Docx4J.FLAG_EXPORT_PREFER_XSL);
            return pdfOut.toByteArray();
        } catch (Exception e) {
            throw new IOException("Failed to export PDF via docx4j: " + e.getMessage(), e);
        }
    }

    // =====================================================================================
    // Map-based paragraph replacement
    // =====================================================================================

    /**
     * Expand loop markers in paragraph blocks (outside tables) across the document,
     * including headers and footers.
     */
    private void processParagraphLoopsWithJsonPath(WordprocessingMLPackage pkg, DocumentContext jsonContext) {
        try {
            processParagraphLoopsInContent(pkg.getMainDocumentPart().getContent(), jsonContext);

            List<SectionWrapper> sections = pkg.getDocumentModel().getSections();
            for (SectionWrapper sw : sections) {
                HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
                if (hfp == null) continue;

                HeaderPart dh = hfp.getDefaultHeader();
                if (dh != null) processParagraphLoopsInContent(dh.getContent(), jsonContext);
                HeaderPart fh = hfp.getFirstHeader();
                if (fh != null) processParagraphLoopsInContent(fh.getContent(), jsonContext);
                HeaderPart eh = hfp.getEvenHeader();
                if (eh != null) processParagraphLoopsInContent(eh.getContent(), jsonContext);

                FooterPart df = hfp.getDefaultFooter();
                if (df != null) processParagraphLoopsInContent(df.getContent(), jsonContext);
                FooterPart ff = hfp.getFirstFooter();
                if (ff != null) processParagraphLoopsInContent(ff.getContent(), jsonContext);
                FooterPart ef = hfp.getEvenFooter();
                if (ef != null) processParagraphLoopsInContent(ef.getContent(), jsonContext);
            }
        } catch (Exception e) {
            System.out.println("WARN: Failed to expand paragraph loops: " + e.getMessage());
        }
    }

    /**
     * Expand loop markers in a given content list (outside tables).
     *
     * Syntax:
     * {{#loop:$arrayPath}}
     *   ... one or more paragraphs ...
     * {{#/loop}}
     */
    @SuppressWarnings("unchecked")
    private void processParagraphLoopsInContent(List<Object> content, DocumentContext jsonContext) {
        if (content == null || content.isEmpty()) return;

        // First, recurse into non-table nested content.
        for (Object o : new ArrayList<>(content)) {
            Object u = unwrap(o);
            if (u instanceof ContentAccessor ca && !(u instanceof Tbl)) {
                processParagraphLoopsInContent(ca.getContent(), jsonContext);
            }
        }

        // Then process only top-level paragraph blocks inside this list.
        for (int endIdx = content.size() - 1; endIdx >= 0; endIdx--) {
            Object endObj = unwrap(content.get(endIdx));
            if (!(endObj instanceof P endP)) continue;

            String endText = getParagraphText(endP).fullText;
            if (!LOOP_END_PATTERN.matcher(endText).find()) continue;

            int startIdx = -1;
            String arrayName = null;
            for (int i = endIdx; i >= 0; i--) {
                Object cand = unwrap(content.get(i));
                if (!(cand instanceof P p)) continue;
                Matcher m = LOOP_START_PATTERN.matcher(getParagraphText(p).fullText);
                if (m.find()) {
                    startIdx = i;
                    arrayName = extractVariableName(m.group(1));
                    break;
                }
            }

            if (startIdx < 0 || arrayName == null || arrayName.isBlank()) {
                removePatternFromParagraph(endP, LOOP_END_PATTERN);
                continue;
            }

            List<Object> templateBlock = new ArrayList<>();
            for (int i = startIdx + 1; i <= endIdx - 1; i++) {
                templateBlock.add(content.get(i));
            }

            List<Object> items;
            try {
                Object read = jsonContext.read("$." + arrayName);
                items = (read instanceof List) ? (List<Object>) read : new ArrayList<>();
            } catch (Exception e) {
                System.out.println("WARN: Loop array not found/readable for paragraph loop: $" + arrayName + " - " + e.getMessage());
                items = new ArrayList<>();
            }

            // Remove markers and template block
            for (int i = endIdx; i >= startIdx; i--) {
                content.remove(i);
            }

            int insertAt = startIdx;
            if (!items.isEmpty() && !templateBlock.isEmpty()) {
                for (int itemIndex = 0; itemIndex < items.size(); itemIndex++) {
                    for (Object blockObj : templateBlock) {
                        Object cloned = XmlUtils.deepCopy(blockObj);
                        Object unwrapped = unwrap(cloned);
                        if (unwrapped instanceof P p) {
                            processAllPlaceholdersWithJsonPath(p, jsonContext, arrayName, itemIndex, jsonContext, items);
                        }
                        content.add(insertAt++, cloned);
                    }
                }
            }

            endIdx = startIdx - 1;
        }
    }

    private void processParagraphWithMap(P p, Map<String, Object> data) {
        removePatternFromParagraph(p, LOOP_START_PATTERN);
        removePatternFromParagraph(p, LOOP_END_PATTERN);

        ParagraphText pt = getParagraphText(p);
        if (pt.fullText.isEmpty()) return;

        Map<String, String> replacements = new HashMap<>();

        Matcher m = PLACEHOLDER_PATTERN.matcher(pt.fullText);
        while (m.find()) {
            String varPath = extractVariableName(m.group(1));
            Object v = data != null ? data.get(varPath) : null;
            replacements.put(m.group(0), v != null ? v.toString() : "");
        }

        Matcher cb = CHECKBOX_PATTERN.matcher(pt.fullText);
        while (cb.find()) {
            String group = extractVariableName(cb.group(1));
            String expected = cb.group(2);
            Object actual = data != null ? data.get(group) : null;
            boolean checked = actual != null && expected.equalsIgnoreCase(actual.toString());
            replacements.put(cb.group(0), checked ? "☒" : "☐");
        }

        Matcher rb = RADIO_PATTERN.matcher(pt.fullText);
        while (rb.find()) {
            String group = extractVariableName(rb.group(1));
            String expected = rb.group(2);
            Object actual = data != null ? data.get(group) : null;
            boolean selected = actual != null && expected.equalsIgnoreCase(actual.toString());
            replacements.put(rb.group(0), selected ? "◉" : "○");
        }

        applyReplacements(p, replacements);
    }

    // =====================================================================================
    // JSONPath paragraph replacement
    // =====================================================================================

    private void processAllPlaceholdersWithJsonPath(P paragraph,
                                                    DocumentContext jsonContext,
                                                    String loopArrayName,
                                                    int currentIndex,
                                                    DocumentContext originalJsonContext,
                                                    List<Object> itemsList) {

        removePatternFromParagraph(paragraph, LOOP_START_PATTERN);
        removePatternFromParagraph(paragraph, LOOP_END_PATTERN);

        String paragraphText = getParagraphText(paragraph).fullText;
        if (paragraphText == null || paragraphText.isEmpty()) return;

        Matcher radioMatcher = RADIO_PATTERN.matcher(paragraphText);
        Matcher checkboxMatcher = CHECKBOX_PATTERN.matcher(paragraphText);
        Matcher regularMatcher = PLACEHOLDER_PATTERN.matcher(paragraphText);
        Matcher loopItemMatcher = LOOP_ITEM_PATTERN.matcher(paragraphText);

        boolean hasRadios = radioMatcher.find();
        boolean hasCheckboxes = checkboxMatcher.find();
        boolean hasRegularPlaceholders = regularMatcher.find();
        boolean hasLoopItems = loopItemMatcher.find();
        if (!hasRadios && !hasCheckboxes && !hasRegularPlaceholders && !hasLoopItems) return;

        Map<String, String> replacements = new HashMap<>();

        if (hasLoopItems) {
            loopItemMatcher.reset();
            while (loopItemMatcher.find()) {
                String arrayName = loopItemMatcher.group(1).trim();
                String indexStr = loopItemMatcher.group(2).trim();
                String property = loopItemMatcher.group(3).trim();
                String fullPlaceholder = loopItemMatcher.group(0);

                if (originalJsonContext == null) {
                    replacements.put(fullPlaceholder, "");
                    continue;
                }

                int targetIndex;
                if (indexStr.equalsIgnoreCase("i")) {
                    if (currentIndex >= 0) targetIndex = currentIndex;
                    else {
                        // Not in a loop context; do NOT blank this placeholder here.
                        // The row-loop processor will handle it when currentIndex is known.
                        continue;
                    }
                } else {
                    try {
                        targetIndex = Integer.parseInt(indexStr);
                    } catch (NumberFormatException e) {
                        // Invalid index; safest is to blank it
                        replacements.put(fullPlaceholder, "");
                        continue;
                    }
                }

                if (itemsList != null && (targetIndex < 0 || targetIndex >= itemsList.size())) {
                    replacements.put(fullPlaceholder, "");
                    continue;
                }

                String jsonPathExpr = "$." + arrayName + "[" + targetIndex + "]." + property;
                try {
                    Object value = originalJsonContext.read(jsonPathExpr);
                    replacements.put(fullPlaceholder, value != null ? value.toString() : "");
                } catch (PathNotFoundException e) {
                    System.out.println("WARN: JSON-Path not found: " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - Replacing with empty string");
                    replacements.put(fullPlaceholder, "");
                } catch (Exception e) {
                    System.out.println("ERROR: Error reading JSON-Path " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - " + e.getMessage() + " - Replacing with empty string");
                    replacements.put(fullPlaceholder, "");
                }
            }
        }

        if (hasRadios) {
            radioMatcher.reset();
            while (radioMatcher.find()) {
                String groupPath = radioMatcher.group(1);
                String expected = radioMatcher.group(2);
                String fullPlaceholder = radioMatcher.group(0);

                String group = extractVariableName(groupPath);
                String jsonPathExpr = "$." + group;
                try {
                    Object actual = jsonContext.read(jsonPathExpr);
                    boolean selected = actual != null && expected.equalsIgnoreCase(actual.toString());
                    replacements.put(fullPlaceholder, selected ? "◉" : "○");
                } catch (Exception e) {
                    replacements.put(fullPlaceholder, "○");
                }
            }
        }

        if (hasCheckboxes) {
            checkboxMatcher.reset();
            while (checkboxMatcher.find()) {
                String groupPath = checkboxMatcher.group(1);
                String expected = checkboxMatcher.group(2);
                String fullPlaceholder = checkboxMatcher.group(0);

                String group = extractVariableName(groupPath);
                String jsonPathExpr = "$." + group;
                try {
                    Object actual = jsonContext.read(jsonPathExpr);
                    boolean checked = actual != null && expected.equalsIgnoreCase(actual.toString());
                    replacements.put(fullPlaceholder, checked ? "☒" : "☐");
                } catch (Exception e) {
                    replacements.put(fullPlaceholder, "☐");
                }
            }
        }

        if (hasRegularPlaceholders) {
            regularMatcher.reset();
            while (regularMatcher.find()) {
                String variablePath = regularMatcher.group(1);
                String fullPlaceholder = regularMatcher.group(0);
                if (replacements.containsKey(fullPlaceholder)) continue;

                String variableName = extractVariableName(variablePath);

                // If user wrote a loop-item placeholder using {{$array[i].prop}} but it wasn't
                // matched by LOOP_ITEM_PATTERN (run split/whitespace edge cases), normalize it here.
                // JSONPath does NOT accept [i], so convert it to a numeric index when we have context.
                if (variableName != null && variableName.contains("[i]")) {
                    if (currentIndex >= 0) {
                        variableName = variableName.replace("[i]", "[" + currentIndex + "]");
                    } else {
                        // No loop context: do not wipe it here; it should be resolved inside loop processing.
                        continue;
                    }
                }

                String jsonPathExpr = convertToJsonPath(variableName);
                try {
                    Object value = jsonContext.read(jsonPathExpr);
                    replacements.put(fullPlaceholder, value != null ? value.toString() : "");
                } catch (PathNotFoundException e) {
                    System.out.println("WARN: JSON-Path not found: " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - Replacing with empty string");
                    replacements.put(fullPlaceholder, "");
                } catch (Exception e) {
                    System.out.println("ERROR: Error reading JSON-Path " + jsonPathExpr + " for placeholder: " + fullPlaceholder + " - " + e.getMessage() + " - Replacing with empty string");
                    replacements.put(fullPlaceholder, "");
                }
            }
        }

        applyReplacements(paragraph, replacements);
    }

    // =====================================================================================
    // Table loops
    // =====================================================================================

    @SuppressWarnings("unchecked")
    private void processTableLoopsWithJsonPath(Tbl tbl, DocumentContext jsonContext) {
        List<Object> content = tbl.getContent();
        if (content == null || content.isEmpty()) return;

        List<Integer> trContentIdx = new ArrayList<>();
        List<Tr> trs = new ArrayList<>();
        for (int i = 0; i < content.size(); i++) {
            Object u = unwrap(content.get(i));
            if (u instanceof Tr tr) {
                trContentIdx.add(i);
                trs.add(tr);
            }
        }
        if (trs.isEmpty()) return;

        for (int endRow = trs.size() - 1; endRow >= 0; endRow--) {
            if (!LOOP_END_PATTERN.matcher(getRowText(trs.get(endRow))).find()) continue;

            int startRow = -1;
            String arrayName = null;
            for (int i = endRow; i >= 0; i--) {
                Matcher sm = LOOP_START_PATTERN.matcher(getRowText(trs.get(i)));
                if (sm.find()) {
                    arrayName = extractVariableName(sm.group(1));
                    startRow = i;
                    break;
                }
            }
            if (startRow < 0 || arrayName == null || arrayName.isEmpty()) continue;

            List<Object> items;
            try {
                Object listObj = jsonContext.read("$." + arrayName);
                if (!(listObj instanceof List)) {
                    removeLoopMarkersFromRowBlock(trs, startRow, endRow);
                    continue;
                }
                items = (List<Object>) listObj;
            } catch (Exception e) {
                removeLoopMarkersFromRowBlock(trs, startRow, endRow);
                continue;
            }

            List<Tr> blockTrs = new ArrayList<>();
            List<Tr> templateCopies = new ArrayList<>();
            for (int i = startRow; i <= endRow; i++) {
                Tr tr = trs.get(i);
                blockTrs.add(tr);
                templateCopies.add((Tr) XmlUtils.deepCopy(tr));
            }

            if (items.isEmpty()) {
                removeTrRange(tbl, trContentIdx.get(startRow), trContentIdx.get(endRow));
            } else {
                for (Tr tr : blockTrs) {
                    processRowForItem(tr, jsonContext, arrayName, 0, jsonContext, items);
                }

                int insertAt = trContentIdx.get(endRow) + 1;
                for (int itemIndex = 1; itemIndex < items.size(); itemIndex++) {
                    for (Tr templateTr : templateCopies) {
                        Tr newTr = (Tr) XmlUtils.deepCopy(templateTr);
                        processRowForItem(newTr, jsonContext, arrayName, itemIndex, jsonContext, items);
                        // Add as JAXBElement to ensure docx4j marshals it correctly
                        content.add(insertAt, new JAXBElement<>(QNAME_TR, Tr.class, newTr));
                        insertAt++;
                    }
                }
            }

            // Re-index after modification
            content = tbl.getContent();
            trContentIdx.clear();
            trs.clear();
            for (int i = 0; i < content.size(); i++) {
                Object u = unwrap(content.get(i));
                if (u instanceof Tr tr) {
                    trContentIdx.add(i);
                    trs.add(tr);
                }
            }

            endRow = startRow - 1;
        }
    }

    private void processRowForItem(Tr tr,
                                   DocumentContext jsonContext,
                                   String loopArrayName,
                                   int currentIndex,
                                   DocumentContext originalJsonContext,
                                   List<Object> itemsList) {

        forEachParagraphInTr(tr, p -> {
            removePatternFromParagraph(p, LOOP_START_PATTERN);
            removePatternFromParagraph(p, LOOP_END_PATTERN);
        });
        forEachParagraphInTr(tr, p -> processAllPlaceholdersWithJsonPath(p, jsonContext, loopArrayName, currentIndex, originalJsonContext, itemsList));
    }

    private void removeLoopMarkersFromRowBlock(List<Tr> trs, int startRow, int endRow) {
        for (int i = startRow; i <= endRow; i++) {
            Tr tr = trs.get(i);
            forEachParagraphInTr(tr, p -> {
                removePatternFromParagraph(p, LOOP_START_PATTERN);
                removePatternFromParagraph(p, LOOP_END_PATTERN);
            });
        }
    }

    private void removeTrRange(Tbl tbl, int startContentIndex, int endContentIndexInclusive) {
        List<Object> content = tbl.getContent();
        for (int i = endContentIndexInclusive; i >= startContentIndex; i--) {
            content.remove(i);
        }
    }

    private String getRowText(Tr tr) {
        StringBuilder sb = new StringBuilder();
        forEachParagraphInTr(tr, p -> sb.append(getParagraphText(p).fullText));
        return sb.toString();
    }

    // =====================================================================================
    // docx4j traversal
    // =====================================================================================

    private WordprocessingMLPackage loadPackage(InputStream in) throws IOException {
        try {
            return WordprocessingMLPackage.load(in);
        } catch (Docx4JException e) {
            throw new IOException("Failed to load DOCX: " + e.getMessage(), e);
        }
    }

    private byte[] savePackage(WordprocessingMLPackage pkg) throws IOException {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            pkg.save(baos);
            return baos.toByteArray();
        } catch (Docx4JException e) {
            throw new IOException("Failed to save DOCX: " + e.getMessage(), e);
        }
    }

    @FunctionalInterface
    private interface ParagraphConsumer { void accept(P p) throws Exception; }

    @FunctionalInterface
    private interface TableConsumer { void accept(Tbl tbl) throws Exception; }

    private void forEachParagraph(WordprocessingMLPackage pkg, ParagraphConsumer consumer) throws IOException {
        try {
            forEachParagraphInContent(pkg.getMainDocumentPart().getContent(), consumer);

            List<SectionWrapper> sections = pkg.getDocumentModel().getSections();
            for (SectionWrapper sw : sections) {
                HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
                if (hfp == null) continue;

                HeaderPart dh = hfp.getDefaultHeader();
                if (dh != null) forEachParagraphInContent(dh.getContent(), consumer);
                HeaderPart fh = hfp.getFirstHeader();
                if (fh != null) forEachParagraphInContent(fh.getContent(), consumer);
                HeaderPart eh = hfp.getEvenHeader();
                if (eh != null) forEachParagraphInContent(eh.getContent(), consumer);

                FooterPart df = hfp.getDefaultFooter();
                if (df != null) forEachParagraphInContent(df.getContent(), consumer);
                FooterPart ff = hfp.getFirstFooter();
                if (ff != null) forEachParagraphInContent(ff.getContent(), consumer);
                FooterPart ef = hfp.getEvenFooter();
                if (ef != null) forEachParagraphInContent(ef.getContent(), consumer);
            }
        } catch (Exception e) {
            throw new IOException("Failed to process paragraphs: " + e.getMessage(), e);
        }
    }

    private void forEachTable(WordprocessingMLPackage pkg, TableConsumer consumer) throws IOException {
        try {
            forEachTableInContent(pkg.getMainDocumentPart().getContent(), consumer);

            List<SectionWrapper> sections = pkg.getDocumentModel().getSections();
            for (SectionWrapper sw : sections) {
                HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
                if (hfp == null) continue;

                HeaderPart dh = hfp.getDefaultHeader();
                if (dh != null) forEachTableInContent(dh.getContent(), consumer);
                HeaderPart fh = hfp.getFirstHeader();
                if (fh != null) forEachTableInContent(fh.getContent(), consumer);
                HeaderPart eh = hfp.getEvenHeader();
                if (eh != null) forEachTableInContent(eh.getContent(), consumer);

                FooterPart df = hfp.getDefaultFooter();
                if (df != null) forEachTableInContent(df.getContent(), consumer);
                FooterPart ff = hfp.getFirstFooter();
                if (ff != null) forEachTableInContent(ff.getContent(), consumer);
                FooterPart ef = hfp.getEvenFooter();
                if (ef != null) forEachTableInContent(ef.getContent(), consumer);
            }
        } catch (Exception e) {
            throw new IOException("Failed to process tables: " + e.getMessage(), e);
        }
    }

    private void forEachParagraphInContent(List<Object> content, ParagraphConsumer consumer) throws Exception {
        if (content == null) return;
        for (Object o : content) {
            Object u = unwrap(o);
            if (u instanceof P p) consumer.accept(p);
            else if (u instanceof ContentAccessor ca) forEachParagraphInContent(ca.getContent(), consumer);
        }
    }

    private void forEachTableInContent(List<Object> content, TableConsumer consumer) throws Exception {
        if (content == null) return;
        for (Object o : content) {
            Object u = unwrap(o);
            if (u instanceof Tbl tbl) consumer.accept(tbl);
            else if (u instanceof ContentAccessor ca) forEachTableInContent(ca.getContent(), consumer);
        }
    }

    private void forEachParagraphInTr(Tr tr, ParagraphConsumer consumer) {
        try { forEachParagraphInContent(tr.getContent(), consumer); } catch (Exception ignored) {}
    }

    private Object unwrap(Object o) {
        if (o instanceof JAXBElement<?> je) return je.getValue();
        return o;
    }

    // =====================================================================================
    // Paragraph text extraction & replacement
    // =====================================================================================

    private static class TextNodeRef {
        final Text text;
        final int start;
        final int end;
        TextNodeRef(Text text, int start, int end) { this.text = text; this.start = start; this.end = end; }
    }

    private static class ParagraphText {
        final String fullText;
        final List<TextNodeRef> nodes;
        ParagraphText(String fullText, List<TextNodeRef> nodes) { this.fullText = fullText; this.nodes = nodes; }
    }

    private ParagraphText getParagraphText(P p) {
        List<Text> texts = new ArrayList<>();
        collectTextNodes(p, texts);
        StringBuilder sb = new StringBuilder();
        List<TextNodeRef> refs = new ArrayList<>();
        int pos = 0;
        for (Text t : texts) {
            String v = t.getValue();
            if (v == null) v = "";
            int start = pos;
            sb.append(v);
            pos += v.length();
            refs.add(new TextNodeRef(t, start, pos));
        }
        return new ParagraphText(sb.toString(), refs);
    }

    private void collectTextNodes(Object root, List<Text> out) {
        Object u = unwrap(root);
        if (u == null) return;
        if (u instanceof Text t) { out.add(t); return; }
        if (u instanceof ContentAccessor ca) {
            for (Object child : ca.getContent()) collectTextNodes(child, out);
        }
    }

    private void replaceInParagraph(P p, int matchStart, int matchEnd, String replacement) {
        ParagraphText pt = getParagraphText(p);
        if (matchStart < 0 || matchEnd > pt.fullText.length() || matchStart >= matchEnd) return;

        int firstIdx = -1, lastIdx = -1;
        for (int i = 0; i < pt.nodes.size(); i++) {
            TextNodeRef ref = pt.nodes.get(i);
            if (firstIdx == -1 && matchStart >= ref.start && matchStart < ref.end) firstIdx = i;
            if (matchEnd > ref.start && matchEnd <= ref.end) { lastIdx = i; break; }
        }
        if (firstIdx == -1) return;
        if (lastIdx == -1) lastIdx = firstIdx;

        TextNodeRef first = pt.nodes.get(firstIdx);
        TextNodeRef last = pt.nodes.get(lastIdx);

        String firstVal = safe(first.text.getValue());
        String lastVal = safe(last.text.getValue());

        int startInFirst = matchStart - first.start;
        int endInLast = matchEnd - last.start;

        String prefix = firstVal.substring(0, clamp(startInFirst, 0, firstVal.length()));
        String suffix = lastVal.substring(clamp(endInLast, 0, lastVal.length()));

        first.text.setValue(prefix + safe(replacement) + suffix);
        for (int i = firstIdx + 1; i <= lastIdx; i++) pt.nodes.get(i).text.setValue("");
    }

    private void removePatternFromParagraph(P p, Pattern pattern) {
        ParagraphText pt = getParagraphText(p);
        Matcher m = pattern.matcher(pt.fullText);
        List<int[]> ranges = new ArrayList<>();
        while (m.find()) ranges.add(new int[]{m.start(), m.end()});
        ranges.sort((a, b) -> Integer.compare(b[0], a[0]));
        for (int[] r : ranges) replaceInParagraph(p, r[0], r[1], "");
    }

    private void applyReplacements(P p, Map<String, String> replacements) {
        if (replacements == null || replacements.isEmpty()) return;
        String current = getParagraphText(p).fullText;
        if (current == null || current.isEmpty()) return;

        class Occ { final int start, end; final String rep; Occ(int s,int e,String r){start=s;end=e;rep=r;} }
        List<Occ> occs = new ArrayList<>();
        for (Map.Entry<String, String> e : replacements.entrySet()) {
            String needle = e.getKey();
            String rep = e.getValue() == null ? "" : e.getValue();
            int idx = 0;
            while (true) {
                int pos = current.indexOf(needle, idx);
                if (pos < 0) break;
                occs.add(new Occ(pos, pos + needle.length(), rep));
                idx = pos + 1;
            }
        }

        occs.sort((a, b) -> Integer.compare(b.start, a.start));
        for (Occ o : occs) {
            replaceInParagraph(p, o.start, o.end, o.rep);
            current = getParagraphText(p).fullText;
        }
    }

    // =====================================================================================
    // Utility
    // =====================================================================================

    private String extractVariableName(String variablePath) {
        if (variablePath == null) return null;
        String trimmed = variablePath.trim();
        if (trimmed.startsWith("$")) return trimmed.substring(1);
        return trimmed;
    }

    private String convertToJsonPath(String variablePath) {
        if (variablePath == null || variablePath.isEmpty()) return "$";
        if (variablePath.startsWith("$")) return variablePath;
        return "$." + variablePath;
    }

    private String safe(String s) { return s == null ? "" : s; }

    private int clamp(int v, int min, int max) { return Math.max(min, Math.min(max, v)); }
}

