package com.example.docxprocessor.service;

import com.example.docxprocessor.model.TemplateValidationIssue;
import com.example.docxprocessor.model.TemplateValidationLoop;
import com.example.docxprocessor.model.TemplateValidationPlaceholder;
import com.example.docxprocessor.model.TemplateValidationReport;
import com.jayway.jsonpath.DocumentContext;
import com.jayway.jsonpath.JsonPath;
import com.jayway.jsonpath.PathNotFoundException;
import jakarta.xml.bind.JAXBElement;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.XmlUtils;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
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
import java.util.LinkedHashMap;
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
    private static final Pattern ANY_TAG_PATTERN = Pattern.compile("\\{\\{[^}]+\\}\\}");

    // More widely-supported Unicode symbols for PDF rendering (FOP font coverage varies).
    private static final String CHECKBOX_CHECKED = "☑"; // U+2611
    private static final String CHECKBOX_UNCHECKED = "☐"; // U+2610
    private static final String RADIO_SELECTED = "●"; // U+25CF
    private static final String RADIO_UNSELECTED = "○"; // U+25CB

    /**
     * Improve PDF fidelity by configuring docx4j font mapping.
     *
     * docx4j-export-fo (Apache FOP) will substitute fonts if it can't find the
     * physical fonts used in the DOCX. That often breaks symbol glyphs (checkbox/radio).
     *
     * This maps document fonts to installed system fonts when possible.
     */
    private static void configureFontsForPdf(WordprocessingMLPackage pkg) {
        try {
            // Limit discovery to common fonts to avoid slow full scans.
            // Add symbol fonts commonly used for checkbox/radio glyphs.
            String regex = "(?i).*(calibri|arial|times|helvetica|courier|dejavu|noto|symbola|segoe ui symbol|apple symbols|wingdings|webdings|symbol).*";
            PhysicalFonts.setRegex(regex);
            PhysicalFonts.discoverPhysicalFonts();

            Mapper fontMapper = new IdentityPlusMapper();
            pkg.setFontMapper(fontMapper);
        } catch (Throwable t) {
            // Don't fail generation; worst-case PDF uses fallback fonts.
            System.out.println("WARN: Font mapping configuration failed; PDF may have reduced fidelity: " + t.getMessage());
        }
    }

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
            configureFontsForPdf(pkg);
            foSettings.setApacheFopMime(FOSettings.MIME_PDF);
            Docx4J.toFO(foSettings, pdfOut, Docx4J.FLAG_EXPORT_PREFER_XSL);
            return pdfOut.toByteArray();
        } catch (Exception e) {
            throw new IOException("Failed to export PDF via docx4j: " + e.getMessage(), e);
        }
    }

    // =====================================================================================
    // Validation (template syntax + JSONPath)
    // =====================================================================================

    public TemplateValidationReport validateTemplateWithJsonPath(InputStream templateInputStream, String jsonSample) throws IOException {
        if (jsonSample == null) jsonSample = "";
        DocumentContext ctx;
        try {
            ctx = JsonPath.parse(jsonSample);
        } catch (Exception e) {
            throw new IOException("Invalid JSON sample: " + e.getMessage(), e);
        }
        return validateTemplateWithJsonPath(templateInputStream, ctx);
    }

    public TemplateValidationReport validateTemplateWithJsonPath(InputStream templateInputStream, DocumentContext jsonContext) throws IOException {
        WordprocessingMLPackage pkg = loadPackage(templateInputStream);

        TemplateValidationReport report = new TemplateValidationReport();
        List<TemplateValidationIssue> issues = new ArrayList<>();
        List<TemplateValidationLoop> loops = new ArrayList<>();
        Map<String, TemplateValidationPlaceholder> placeholdersByKey = new HashMap<>();

        // Validate paragraph loops (outside tables) first, to validate loop-item placeholders with [i].
        validateParagraphLoops(pkg, jsonContext, loops, issues, placeholdersByKey);

        // Validate table loops.
        forEachTable(pkg, tbl -> validateTableLoops(tbl, jsonContext, loops, issues, placeholdersByKey));

        // Global scan of all paragraphs for placeholders/tags (includes header/footer).
        forEachParagraph(pkg, p -> scanTextForTags(getParagraphText(p).fullText, jsonContext, null, -1, placeholdersByKey, issues));

        // Deduplicate issues. The same placeholder can be scanned multiple times:
        // - once in table/loop validation block scans
        // - again in the global paragraph scan (which also traverses inside tables)
        // This can produce duplicate issue objects with identical content.
        List<TemplateValidationIssue> dedupedIssues = dedupeIssues(issues);

        // Finalize
        report.setLoops(loops);
        report.setIssues(dedupedIssues);
        report.setPlaceholders(new ArrayList<>(placeholdersByKey.values()));

        int errors = 0;
        int warns = 0;
        for (TemplateValidationIssue i : dedupedIssues) {
            if ("ERROR".equalsIgnoreCase(i.getSeverity())) errors++;
            else if ("WARN".equalsIgnoreCase(i.getSeverity())) warns++;
        }
        report.setErrorCount(errors);
        report.setWarningCount(warns);
        report.setValid(errors == 0);
        return report;
    }

    private List<TemplateValidationIssue> dedupeIssues(List<TemplateValidationIssue> issues) {
        if (issues == null || issues.isEmpty()) return issues;
        LinkedHashMap<String, TemplateValidationIssue> uniq = new LinkedHashMap<>();
        for (TemplateValidationIssue i : issues) {
            if (i == null) continue;
            String key =
                    safe(i.getSeverity()) + "|" +
                    safe(i.getCode()) + "|" +
                    safe(i.getMessage()) + "|" +
                    safe(i.getContext());
            uniq.putIfAbsent(key, i);
        }
        return new ArrayList<>(uniq.values());
    }

    private void validateParagraphLoops(WordprocessingMLPackage pkg,
                                        DocumentContext jsonContext,
                                        List<TemplateValidationLoop> loops,
                                        List<TemplateValidationIssue> issues,
                                        Map<String, TemplateValidationPlaceholder> placeholdersByKey) {
        try {
            validateParagraphLoopsInContent(pkg.getMainDocumentPart().getContent(), jsonContext, loops, issues, placeholdersByKey);

            List<SectionWrapper> sections = pkg.getDocumentModel().getSections();
            for (SectionWrapper sw : sections) {
                HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
                if (hfp == null) continue;

                HeaderPart dh = hfp.getDefaultHeader();
                if (dh != null) validateParagraphLoopsInContent(dh.getContent(), jsonContext, loops, issues, placeholdersByKey);
                HeaderPart fh = hfp.getFirstHeader();
                if (fh != null) validateParagraphLoopsInContent(fh.getContent(), jsonContext, loops, issues, placeholdersByKey);
                HeaderPart eh = hfp.getEvenHeader();
                if (eh != null) validateParagraphLoopsInContent(eh.getContent(), jsonContext, loops, issues, placeholdersByKey);

                FooterPart df = hfp.getDefaultFooter();
                if (df != null) validateParagraphLoopsInContent(df.getContent(), jsonContext, loops, issues, placeholdersByKey);
                FooterPart ff = hfp.getFirstFooter();
                if (ff != null) validateParagraphLoopsInContent(ff.getContent(), jsonContext, loops, issues, placeholdersByKey);
                FooterPart ef = hfp.getEvenFooter();
                if (ef != null) validateParagraphLoopsInContent(ef.getContent(), jsonContext, loops, issues, placeholdersByKey);
            }
        } catch (Exception e) {
            issues.add(new TemplateValidationIssue("WARN", "VALIDATION_FAILED", "Failed to validate paragraph loops: " + e.getMessage(), null));
        }
    }

    @SuppressWarnings("unchecked")
    private void validateParagraphLoopsInContent(List<Object> content,
                                                 DocumentContext jsonContext,
                                                 List<TemplateValidationLoop> loops,
                                                 List<TemplateValidationIssue> issues,
                                                 Map<String, TemplateValidationPlaceholder> placeholdersByKey) {
        if (content == null || content.isEmpty()) return;

        // Recurse into non-table nested content.
        for (Object o : new ArrayList<>(content)) {
            Object u = unwrap(o);
            if (u instanceof ContentAccessor ca && !(u instanceof Tbl)) {
                validateParagraphLoopsInContent(ca.getContent(), jsonContext, loops, issues, placeholdersByKey);
            }
        }

        List<Integer> pContentIdx = new ArrayList<>();
        List<P> ps = new ArrayList<>();
        for (int i = 0; i < content.size(); i++) {
            Object u = unwrap(content.get(i));
            if (u instanceof P p) {
                pContentIdx.add(i);
                ps.add(p);
            }
        }
        if (ps.isEmpty()) return;

        boolean[] startMatched = new boolean[ps.size()];
        for (int endP = ps.size() - 1; endP >= 0; endP--) {
            String endText = getParagraphText(ps.get(endP)).fullText;
            if (!LOOP_END_PATTERN.matcher(endText).find()) continue;

            int startP = -1;
            String arrayName = null;
            for (int i = endP; i >= 0; i--) {
                Matcher sm = LOOP_START_PATTERN.matcher(getParagraphText(ps.get(i)).fullText);
                if (sm.find()) {
                    startP = i;
                    arrayName = extractVariableName(sm.group(1));
                    break;
                }
            }

            if (startP < 0 || arrayName == null || arrayName.isBlank()) {
                issues.add(new TemplateValidationIssue("ERROR", "LOOP_UNMATCHED_END", "Found {{#/loop}} without matching {{#loop:$...}}", "{{#/loop}}"));
                continue;
            }

            startMatched[startP] = true;
            TemplateValidationLoop loop = new TemplateValidationLoop("PARAGRAPH", arrayName);
            loop.setPaired(true);

            List<Object> items = null;
            try {
                Object read = jsonContext.read("$." + arrayName);
                if (read instanceof List) items = (List<Object>) read;
                else {
                    issues.add(new TemplateValidationIssue("ERROR", "LOOP_ARRAY_NOT_LIST", "Loop array is not a JSON array: $" + arrayName, "{{#loop:$" + arrayName + "}}"));
                }
            } catch (PathNotFoundException e) {
                issues.add(new TemplateValidationIssue("ERROR", "LOOP_ARRAY_NOT_FOUND", "Loop array not found in JSON: $" + arrayName, "{{#loop:$" + arrayName + "}}"));
            } catch (Exception e) {
                issues.add(new TemplateValidationIssue("ERROR", "LOOP_ARRAY_READ_ERROR", "Error reading loop array $" + arrayName + ": " + e.getMessage(), "{{#loop:$" + arrayName + "}}"));
            }

            if (items != null) loop.setArrayLength(items.size());
            loops.add(loop);

            int loopIndexForValidation = (items != null && !items.isEmpty()) ? 0 : -1;

            // Validate tags inside the block (exclusive of marker paragraphs)
            for (int i = startP + 1; i <= endP - 1; i++) {
                scanTextForTags(getParagraphText(ps.get(i)).fullText, jsonContext, arrayName, loopIndexForValidation, placeholdersByKey, issues);
            }

            endP = startP - 1;
        }

        // Unmatched starts
        for (int i = 0; i < ps.size(); i++) {
            String txt = getParagraphText(ps.get(i)).fullText;
            Matcher sm = LOOP_START_PATTERN.matcher(txt);
            if (!sm.find()) continue;
            String arrayName = extractVariableName(sm.group(1));
            if (startMatched[i]) continue;
            TemplateValidationLoop loop = new TemplateValidationLoop("PARAGRAPH", arrayName);
            loop.setPaired(false);
            loops.add(loop);
            issues.add(new TemplateValidationIssue("ERROR", "LOOP_UNMATCHED_START", "Found {{#loop:$...}} without matching {{#/loop}}", sm.group(0)));
        }
    }

    @SuppressWarnings("unchecked")
    private void validateTableLoops(Tbl tbl,
                                    DocumentContext jsonContext,
                                    List<TemplateValidationLoop> loops,
                                    List<TemplateValidationIssue> issues,
                                    Map<String, TemplateValidationPlaceholder> placeholdersByKey) {
        List<Object> content = tbl.getContent();
        if (content == null || content.isEmpty()) return;

        List<Tr> trs = new ArrayList<>();
        for (Object o : content) {
            Object u = unwrap(o);
            if (u instanceof Tr tr) trs.add(tr);
        }
        if (trs.isEmpty()) return;

        boolean[] startMatched = new boolean[trs.size()];

        for (int endRow = trs.size() - 1; endRow >= 0; endRow--) {
            String endText = getRowText(trs.get(endRow));
            if (!LOOP_END_PATTERN.matcher(endText).find()) continue;

            int startRow = -1;
            String arrayName = null;
            for (int i = endRow; i >= 0; i--) {
                Matcher sm = LOOP_START_PATTERN.matcher(getRowText(trs.get(i)));
                if (sm.find()) {
                    startRow = i;
                    arrayName = extractVariableName(sm.group(1));
                    break;
                }
            }
            if (startRow < 0 || arrayName == null || arrayName.isBlank()) {
                issues.add(new TemplateValidationIssue("ERROR", "LOOP_UNMATCHED_END", "Found {{#/loop}} in table without matching {{#loop:$...}}", "{{#/loop}}"));
                continue;
            }

            startMatched[startRow] = true;

            TemplateValidationLoop loop = new TemplateValidationLoop("TABLE", arrayName);
            loop.setPaired(true);

            List<Object> items = null;
            try {
                Object read = jsonContext.read("$." + arrayName);
                if (read instanceof List) items = (List<Object>) read;
                else {
                    issues.add(new TemplateValidationIssue("ERROR", "LOOP_ARRAY_NOT_LIST", "Loop array is not a JSON array: $" + arrayName, "{{#loop:$" + arrayName + "}}"));
                }
            } catch (PathNotFoundException e) {
                issues.add(new TemplateValidationIssue("ERROR", "LOOP_ARRAY_NOT_FOUND", "Loop array not found in JSON: $" + arrayName, "{{#loop:$" + arrayName + "}}"));
            } catch (Exception e) {
                issues.add(new TemplateValidationIssue("ERROR", "LOOP_ARRAY_READ_ERROR", "Error reading loop array $" + arrayName + ": " + e.getMessage(), "{{#loop:$" + arrayName + "}}"));
            }

            if (items != null) loop.setArrayLength(items.size());
            loops.add(loop);

            int loopIndexForValidation = (items != null && !items.isEmpty()) ? 0 : -1;
            for (int i = startRow; i <= endRow; i++) {
                scanTextForTags(getRowText(trs.get(i)), jsonContext, arrayName, loopIndexForValidation, placeholdersByKey, issues);
            }

            endRow = startRow - 1;
        }

        // Unmatched starts in table
        for (int i = 0; i < trs.size(); i++) {
            String txt = getRowText(trs.get(i));
            Matcher sm = LOOP_START_PATTERN.matcher(txt);
            if (!sm.find()) continue;
            String arrayName = extractVariableName(sm.group(1));
            if (startMatched[i]) continue;
            TemplateValidationLoop loop = new TemplateValidationLoop("TABLE", arrayName);
            loop.setPaired(false);
            loops.add(loop);
            issues.add(new TemplateValidationIssue("ERROR", "LOOP_UNMATCHED_START", "Found {{#loop:$...}} in table without matching {{#/loop}}", sm.group(0)));
        }
    }

    private void scanTextForTags(String text,
                                 DocumentContext jsonContext,
                                 String loopArrayName,
                                 int loopIndexForValidation,
                                 Map<String, TemplateValidationPlaceholder> placeholdersByKey,
                                 List<TemplateValidationIssue> issues) {
        if (text == null || text.isEmpty()) return;

        // Known patterns
        extractAndValidateLoopItems(text, jsonContext, loopArrayName, loopIndexForValidation, placeholdersByKey, issues);
        extractAndValidateCheckboxRadio(text, jsonContext, placeholdersByKey, issues);
        extractAndValidateRegularPlaceholders(text, jsonContext, loopIndexForValidation, placeholdersByKey, issues);

        // Unknown tags detection
        Matcher any = ANY_TAG_PATTERN.matcher(text);
        while (any.find()) {
            String raw = any.group(0);
            if (PLACEHOLDER_PATTERN.matcher(raw).matches()) continue;
            if (CHECKBOX_PATTERN.matcher(raw).matches()) continue;
            if (RADIO_PATTERN.matcher(raw).matches()) continue;
            if (LOOP_START_PATTERN.matcher(raw).matches()) continue;
            if (LOOP_END_PATTERN.matcher(raw).matches()) continue;
            // Loop-item placeholders are a subset of PLACEHOLDER_PATTERN, but keep explicit.
            if (LOOP_ITEM_PATTERN.matcher(raw).matches()) continue;

            recordPlaceholder(placeholdersByKey, buildUnknownTag(raw));
            issues.add(new TemplateValidationIssue("WARN", "UNKNOWN_TAG", "Unknown template tag found", raw));
        }
    }

    private void extractAndValidateLoopItems(String text,
                                            DocumentContext jsonContext,
                                            String loopArrayName,
                                            int loopIndexForValidation,
                                            Map<String, TemplateValidationPlaceholder> placeholdersByKey,
                                            List<TemplateValidationIssue> issues) {
        Matcher m = LOOP_ITEM_PATTERN.matcher(text);
        while (m.find()) {
            String arrayName = m.group(1).trim();
            String indexStr = m.group(2).trim();
            String property = m.group(3).trim();
            String raw = m.group(0);

            boolean loopDependent = indexStr.equalsIgnoreCase("i");
            Integer idx = null;
            if (loopDependent) {
                if (loopIndexForValidation >= 0) idx = loopIndexForValidation;
            } else {
                try { idx = Integer.parseInt(indexStr); } catch (NumberFormatException ignored) {}
            }

            String jsonPath;
            if (idx != null) {
                jsonPath = "$." + arrayName + "[" + idx + "]." + property;
            } else {
                // Can't validate i without loop context; attempt with index 0 for basic validation, but mark loopDependent.
                jsonPath = "$." + arrayName + "[0]." + property;
            }

            TemplateValidationPlaceholder p = new TemplateValidationPlaceholder(raw, "LOOP_ITEM");
            p.setLoopDependent(loopDependent);
            validateJsonPath(jsonContext, jsonPath, p, issues);

            if (loopDependent && loopIndexForValidation < 0) {
                // Only warn if this placeholder wasn't already validated in a loop context.
                if (!isPlaceholderAlreadyValidated(placeholdersByKey, raw)) {
                    issues.add(new TemplateValidationIssue("WARN", "LOOP_INDEX_REQUIRED", "Loop-item placeholder uses [i] but no loop context was detected for validation", raw));
                    p.setMessage("Loop-dependent placeholder; validated using index [0] only.");
                }
            }

            recordPlaceholder(placeholdersByKey, p);
        }
    }

    private void extractAndValidateCheckboxRadio(String text,
                                                 DocumentContext jsonContext,
                                                 Map<String, TemplateValidationPlaceholder> placeholdersByKey,
                                                 List<TemplateValidationIssue> issues) {
        Matcher rb = RADIO_PATTERN.matcher(text);
        while (rb.find()) {
            String groupPath = rb.group(1);
            String expected = rb.group(2);
            String raw = rb.group(0);
            String group = extractVariableName(groupPath);
            String jsonPath = "$." + group;
            TemplateValidationPlaceholder p = new TemplateValidationPlaceholder(raw, "RADIO");
            p.setMessage("Expected value: " + expected);
            validateJsonPath(jsonContext, jsonPath, p, issues);
            recordPlaceholder(placeholdersByKey, p);
        }

        Matcher cb = CHECKBOX_PATTERN.matcher(text);
        while (cb.find()) {
            String groupPath = cb.group(1);
            String expected = cb.group(2);
            String raw = cb.group(0);
            String group = extractVariableName(groupPath);
            String jsonPath = "$." + group;
            TemplateValidationPlaceholder p = new TemplateValidationPlaceholder(raw, "CHECKBOX");
            p.setMessage("Expected value: " + expected);
            validateJsonPath(jsonContext, jsonPath, p, issues);
            recordPlaceholder(placeholdersByKey, p);
        }
    }

    private void extractAndValidateRegularPlaceholders(String text,
                                                       DocumentContext jsonContext,
                                                       int loopIndexForValidation,
                                                       Map<String, TemplateValidationPlaceholder> placeholdersByKey,
                                                       List<TemplateValidationIssue> issues) {
        Matcher m = PLACEHOLDER_PATTERN.matcher(text);
        while (m.find()) {
            String varPath = m.group(1);
            String raw = m.group(0);
            String variableName = extractVariableName(varPath);
            boolean loopDependent = variableName != null && variableName.contains("[i]");

            String normalized = variableName;
            if (loopDependent) {
                if (loopIndexForValidation >= 0) normalized = normalized.replace("[i]", "[" + loopIndexForValidation + "]");
                else normalized = normalized.replace("[i]", "[0]");
            }

            String jsonPath = convertToJsonPath(normalized);
            TemplateValidationPlaceholder p = new TemplateValidationPlaceholder(raw, "PLACEHOLDER");
            p.setLoopDependent(loopDependent);
            validateJsonPath(jsonContext, jsonPath, p, issues);

            if (loopDependent && loopIndexForValidation < 0) {
                // Only warn if this placeholder wasn't already validated in a loop context.
                // The loop validation scans placeholders with correct context first, so if
                // a placeholder with [i] already exists in placeholdersByKey, it was likely
                // validated correctly in a loop block.
                if (!isPlaceholderAlreadyValidated(placeholdersByKey, raw)) {
                    issues.add(new TemplateValidationIssue("WARN", "LOOP_INDEX_REQUIRED", "Placeholder uses [i] but no loop context was detected for validation", raw));
                    p.setMessage("Loop-dependent placeholder; validated using index [0] only.");
                }
            }

            recordPlaceholder(placeholdersByKey, p);
        }
    }

    private void validateJsonPath(DocumentContext jsonContext,
                                  String jsonPathExpr,
                                  TemplateValidationPlaceholder p,
                                  List<TemplateValidationIssue> issues) {
        p.setJsonPath(jsonPathExpr);

        boolean syntaxOk;
        try {
            JsonPath.compile(jsonPathExpr);
            syntaxOk = true;
        } catch (RuntimeException e) {
            syntaxOk = false;
            p.setJsonPathSyntaxValid(false);
            p.setJsonPathFound(false);
            issues.add(new TemplateValidationIssue("ERROR", "JSON_PATH_INVALID", "Invalid JSONPath syntax: " + e.getMessage(), p.getRaw()));
            return;
        }

        p.setJsonPathSyntaxValid(syntaxOk);
        try {
            Object v = jsonContext.read(jsonPathExpr);
            p.setJsonPathFound(true);
            if (v == null) {
                issues.add(new TemplateValidationIssue("WARN", "JSON_PATH_NULL", "JSONPath exists but value is null: " + jsonPathExpr, p.getRaw()));
            }
        } catch (PathNotFoundException e) {
            p.setJsonPathFound(false);
            // As requested: include missing JSON paths in the error list.
            issues.add(new TemplateValidationIssue("ERROR", "JSON_PATH_NOT_FOUND", "JSONPath not found in JSON: " + jsonPathExpr, p.getRaw()));
        } catch (Exception e) {
            p.setJsonPathFound(false);
            issues.add(new TemplateValidationIssue("WARN", "JSON_PATH_READ_ERROR", "Error reading JSONPath " + jsonPathExpr + ": " + e.getMessage(), p.getRaw()));
        }
    }

    private TemplateValidationPlaceholder buildUnknownTag(String raw) {
        TemplateValidationPlaceholder p = new TemplateValidationPlaceholder(raw, "UNKNOWN_TAG");
        p.setJsonPath(null);
        p.setJsonPathSyntaxValid(true);
        p.setJsonPathFound(true);
        p.setLoopDependent(false);
        return p;
    }

    /**
     * Check if a placeholder with the same raw text was already validated.
     * Used to avoid false-positive warnings when a placeholder is validated
     * in a loop context first, then scanned again in the global pass.
     */
    private boolean isPlaceholderAlreadyValidated(Map<String, TemplateValidationPlaceholder> placeholdersByKey, String raw) {
        if (raw == null || placeholdersByKey == null || placeholdersByKey.isEmpty()) return false;
        for (TemplateValidationPlaceholder existing : placeholdersByKey.values()) {
            if (raw.equals(existing.getRaw())) {
                return true;
            }
        }
        return false;
    }

    private void recordPlaceholder(Map<String, TemplateValidationPlaceholder> placeholdersByKey,
                                   TemplateValidationPlaceholder p) {
        String key = p.getType() + "|" + safe(p.getRaw()) + "|" + safe(p.getJsonPath());
        TemplateValidationPlaceholder existing = placeholdersByKey.get(key);
        if (existing == null) {
            p.setOccurrences(1);
            placeholdersByKey.put(key, p);
        } else {
            existing.setOccurrences(existing.getOccurrences() + 1);
            // Prefer keeping first message if present; otherwise take the latest.
            if ((existing.getMessage() == null || existing.getMessage().isBlank()) && p.getMessage() != null) {
                existing.setMessage(p.getMessage());
            }
            // Upgrade flags conservatively: if any occurrence is false, keep false.
            existing.setJsonPathSyntaxValid(existing.isJsonPathSyntaxValid() && p.isJsonPathSyntaxValid());
            existing.setJsonPathFound(existing.isJsonPathFound() && p.isJsonPathFound());
            existing.setLoopDependent(existing.isLoopDependent() || p.isLoopDependent());
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
            replacements.put(cb.group(0), checked ? CHECKBOX_CHECKED : CHECKBOX_UNCHECKED);
        }

        Matcher rb = RADIO_PATTERN.matcher(pt.fullText);
        while (rb.find()) {
            String group = extractVariableName(rb.group(1));
            String expected = rb.group(2);
            Object actual = data != null ? data.get(group) : null;
            boolean selected = actual != null && expected.equalsIgnoreCase(actual.toString());
            replacements.put(rb.group(0), selected ? RADIO_SELECTED : RADIO_UNSELECTED);
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
                    replacements.put(fullPlaceholder, selected ? RADIO_SELECTED : RADIO_UNSELECTED);
                } catch (Exception e) {
                    replacements.put(fullPlaceholder, RADIO_UNSELECTED);
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
                    replacements.put(fullPlaceholder, checked ? CHECKBOX_CHECKED : CHECKBOX_UNCHECKED);
                } catch (Exception e) {
                    replacements.put(fullPlaceholder, CHECKBOX_UNCHECKED);
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

