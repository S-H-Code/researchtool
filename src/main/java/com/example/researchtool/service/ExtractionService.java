package com.example.researchtool.service;

import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import technology.tabula.ObjectExtractor;
import technology.tabula.Page;
import technology.tabula.RectangularTextContainer;
import technology.tabula.Table;
import technology.tabula.extractors.BasicExtractionAlgorithm;
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm;

@Service
public class ExtractionService {

    private static final Pattern NUM_PATTERN =
            Pattern.compile("\\(?-?\\d[\\d,]*\\.?\\d*\\)?");

    private static final Pattern FY_PATTERN =
            Pattern.compile("(FY\\s*\\d{2,4})|(20\\d{2})", Pattern.CASE_INSENSITIVE);


    // ================= MAIN METHOD =================
    public ByteArrayInputStream processPdf(MultipartFile file) throws Exception {

        byte[] pdfBytes = file.getBytes();

        LinkedHashMap<String, List<Double>> merged = new LinkedHashMap<>();
        List<String> headers = new ArrayList<>();

        try (PDDocument doc = PDDocument.load(pdfBytes)) {

            List<Integer> candidatePages = findCandidatePages(doc);
            if (candidatePages.isEmpty()) {
                for (int i = 1; i <= doc.getNumberOfPages(); i++) {
                    candidatePages.add(i);
                }
            }

            // -------- TABULA --------
            List<List<List<String>>> tables =
                    extractTablesFromPages(doc, candidatePages);

            for (List<List<String>> table : tables) {
                for (List<String> row : table) {

                    if (row.size() < 2) continue;

                    String label = row.get(0).trim();
                    if (label.isBlank()) continue;

                    List<Double> values = new ArrayList<>();

                    for (int i = 1; i < row.size(); i++) {
                        Double val = parseNumber(row.get(i));
                        values.add(val);
                    }

                    if (!values.isEmpty())
                        merged.putIfAbsent(label, values);
                }
            }
        }

        // -------- OCR FALLBACK --------
        if (merged.isEmpty()) {
            try (PDDocument doc = PDDocument.load(pdfBytes)) {
                for (int i = 0; i < doc.getNumberOfPages(); i++) {
                    String text = ocrPage(doc, i);
                    parseOcrText(text, merged, headers);
                }
            }
        }

        return buildExcel(merged);
    }

    // ================= OCR =================
   private String ocrPage(PDDocument doc, int pageIndex)
        throws IOException {

    PDFRenderer renderer = new PDFRenderer(doc);
    BufferedImage image =
            renderer.renderImageWithDPI(pageIndex, 300, ImageType.RGB);

    File tempImage = File.createTempFile("ocr_page", ".png");
    ImageIO.write(image, "png", tempImage);

    ProcessBuilder pb = new ProcessBuilder(
            "tesseract",
            tempImage.getAbsolutePath(),
            "stdout",
            "-l",
            "eng"
    );

    Process process = pb.start();

    BufferedReader reader =
            new BufferedReader(new InputStreamReader(process.getInputStream()));

    StringBuilder output = new StringBuilder();
    String line;

    while ((line = reader.readLine()) != null) {
        output.append(line).append("\n");
    }

    tempImage.delete();

    return output.toString();
}

    private void parseOcrText(String text,
                              LinkedHashMap<String, List<Double>> merged,
                              List<String> headers) {

        String[] lines = text.split("\\R");

        for (String raw : lines) {

            String line = raw.trim();
            if (line.length() < 5) continue;

            String lower = line.toLowerCase();

            if (headers.isEmpty()) {
                Matcher fyMatcher = FY_PATTERN.matcher(line);
                while (fyMatcher.find()) {
                    headers.add(fyMatcher.group());
                }
                continue;
            }

            if (lower.contains("balance sheet")
                    || lower.contains("cash flow")) {
                break;
            }

            line = line.replace("O", "0");
            line = line.replace("l", "1");

            Matcher m = NUM_PATTERN.matcher(line);
            List<Double> nums = new ArrayList<>();

            while (m.find()) {
                String rawNum = m.group();

                if (rawNum.startsWith("(") && rawNum.endsWith(")")) {
                    rawNum = "-" + rawNum.substring(1, rawNum.length() - 1);
                }

                Double d = parseNumber(rawNum);
                if (d != null) nums.add(d);
            }

            if (nums.isEmpty()) continue;

            Matcher firstMatch = NUM_PATTERN.matcher(line);
            if (!firstMatch.find()) continue;

            int firstIndex = firstMatch.start();
            if (firstIndex <= 0) continue;

            String label = line.substring(0, firstIndex)
                    .replaceAll("[^a-zA-Z()\\-\\s]", "")
                    .replaceAll("\\s{2,}", " ")
                    .trim();

            if (label.length() < 3) continue;

            merged.put(label, nums);
        }
    }

    // ================= TABULA =================
    private List<List<List<String>>> extractTablesFromPages(
            PDDocument doc, List<Integer> pages) throws IOException {

        List<List<List<String>>> out = new ArrayList<>();

        try (ObjectExtractor extractor = new ObjectExtractor(doc)) {

            SpreadsheetExtractionAlgorithm sea =
                    new SpreadsheetExtractionAlgorithm();
            BasicExtractionAlgorithm bea =
                    new BasicExtractionAlgorithm();

            for (int p : pages) {

                Page page = extractor.extract(p);

                List<? extends Table> tables = sea.extract(page);

                if (tables == null || tables.isEmpty()) {
                    tables = bea.extract(page);
                }

                for (Table t : tables) {

                     List<List<String>> rows = new ArrayList<>();

                for (List<RectangularTextContainer> tr : t.getRows()) {
                     List<String> row = new ArrayList<>();

                for (RectangularTextContainer cell : tr) {
                     row.add(cell.getText());
                }

                     rows.add(row);
                }

                    out.add(rows);
                }
            }
        }
        return out;
    }

    // ================= EXCEL BUILDER =================
   private ByteArrayInputStream buildExcel(
        LinkedHashMap<String, List<Double>> merged)
        throws IOException {

    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet("Income Statement");

    DataFormat df = workbook.createDataFormat();

    // ===== STYLES =====
    Font boldFont = workbook.createFont();
    boldFont.setBold(true);

    Font headerFont = workbook.createFont();
    headerFont.setBold(true);
    headerFont.setColor(IndexedColors.WHITE.getIndex());

    CellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFont(headerFont);
    headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    headerStyle.setAlignment(HorizontalAlignment.CENTER);

    CellStyle numberStyle = workbook.createCellStyle();
    numberStyle.setDataFormat(df.getFormat("₹ ##,##,##0.00"));

    CellStyle percentStyle = workbook.createCellStyle();
    percentStyle.setDataFormat(df.getFormat("0.00%"));

    CellStyle totalStyle = workbook.createCellStyle();
    totalStyle.setFont(boldFont);
    totalStyle.setDataFormat(df.getFormat("₹ ##,##,##0.00"));
    totalStyle.setBorderTop(BorderStyle.MEDIUM);

    CellStyle sectionStyle = workbook.createCellStyle();
    sectionStyle.setFont(boldFont);

    int rowIdx = 0;

    // ===== TITLE =====
    Row title = sheet.createRow(rowIdx++);
    Cell titleCell = title.createCell(0);
    titleCell.setCellValue("INCOME STATEMENT");
    titleCell.getCellStyle().setFont(boldFont);

    rowIdx++;

    // ===== DYNAMIC HEADER =====
    // Detect actual year count
int maxYears = 0;
for (List<Double> v : merged.values()) {
    maxYears = Math.max(maxYears, v.size());
}

// Create header row
Row header = sheet.createRow(rowIdx++);
header.createCell(0).setCellValue("Particulars");

// If only 3 years detected, only 3 columns created
for (int i = 0; i < maxYears; i++) {
    header.createCell(i + 1).setCellValue("Year " + (i + 1));
}

    // ===== DATA =====
    for (Map.Entry<String, List<Double>> entry : merged.entrySet()) {

        String label = entry.getKey();
        List<Double> vals = entry.getValue();

        Row row = sheet.createRow(rowIdx++);
        Cell labelCell = row.createCell(0);

        boolean isSection =
                label.equalsIgnoreCase("Expenses")
                        || label.equalsIgnoreCase("Finance Costs")
                        || label.equalsIgnoreCase("Employee Benefit Expenses")
                        || label.equalsIgnoreCase("Other Expenses");

        boolean isTotal =
                label.toLowerCase().contains("total")
                        || label.toLowerCase().contains("profit")
                        || label.toLowerCase().contains("ebitda");

        if (isSection) {
            labelCell.setCellValue(label);
            labelCell.setCellStyle(sectionStyle);
            rowIdx++; // add spacing
            continue;
        }

        if (!isTotal) {
            labelCell.setCellValue("   " + label);
        } else {
            labelCell.setCellValue(label);
            labelCell.setCellStyle(sectionStyle);
        }

        for (int i = 0; i < vals.size(); i++) {

            Double value = vals.get(i);
            if (value == null) continue;

            Cell cell = row.createCell(i + 1);

            if (label.toLowerCase().contains("margin")) {
                cell.setCellValue(value);
                cell.setCellStyle(percentStyle);
            } else if (isTotal) {
                cell.setCellValue(value);
                cell.setCellStyle(totalStyle);
            } else {
                cell.setCellValue(value);
                cell.setCellStyle(numberStyle);
            }
        }
    }

    for (int i = 0; i <= maxYears; i++) {
        sheet.autoSizeColumn(i);
    }

    return writeWorkbook(workbook);
}
    // ================= HELPERS =================
    private Double parseNumber(String s) {
        if (s == null) return null;

        String num = s.replace(",", "")
                .replace("(", "-")
                .replace(")", "");

        try {
            return Double.parseDouble(num);
        } catch (Exception e) {
            return null;
        }
    }

    private List<Integer> findCandidatePages(PDDocument doc)
            throws IOException {

        PDFTextStripper stripper = new PDFTextStripper();
        List<Integer> pages = new ArrayList<>();

        for (int i = 1; i <= doc.getNumberOfPages(); i++) {

            stripper.setStartPage(i);
            stripper.setEndPage(i);

            String text = stripper.getText(doc).toLowerCase();

            if (text.contains("revenue")
                    || text.contains("profit")
                    || text.contains("expenses")) {

                pages.add(i);
            }
        }
        return pages;
    }

    private ByteArrayInputStream writeWorkbook(Workbook workbook)
            throws IOException {

        try (ByteArrayOutputStream out =
                     new ByteArrayOutputStream()) {

            workbook.write(out);
            workbook.close();
            return new ByteArrayInputStream(out.toByteArray());
        }
    }
}