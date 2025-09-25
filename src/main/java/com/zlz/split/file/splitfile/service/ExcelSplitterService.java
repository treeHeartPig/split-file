package com.zlz.split.file.splitfile.service;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.zlz.split.file.splitfile.util.MinioUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

@Service
public class ExcelSplitterService {
    @Autowired
    private MinioUtil minioUtil;

    /**
     * 主方法：处理上传的Excel文件，按Sheet切分
     */
    public Map<String,Object> splitExcelBySheet(MultipartFile file) throws Exception {
        Map<String,Object> result = new HashMap<>();

        Map<Integer,String> sheets = new HashMap<>();
        Map<Integer,String> thumbs = new HashMap<>();

        try (InputStream is = file.getInputStream();
             Workbook workbook = createWorkbook(is, file.getOriginalFilename())) {
            int totalSheets = workbook.getNumberOfSheets();
            result.put("totalPages",totalSheets);
            for (int i = 0; i < totalSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) continue;

                // 将Excel Sheet转换为PDF
                byte[] pdfBytes = convertSheetToPdf(sheet);

                String sheetFileName = String.format("%s_%s.pdf",
                        getBaseName(file.getOriginalFilename()),
                        sheet.getSheetName().replaceAll("[^a-zA-Z0-9\u4e00-\u9fa5]", "_"));

                // 上传到 MinIO
                String path = minioUtil.uploadToMinio(pdfBytes, sheetFileName, "application/pdf");
                sheets.put(i,minioUtil.getFileUrl(path));

                // 生成并上传缩略图
                byte[] thumbnailData = generateThumbnail(sheet, 700, 200);
                String thumbObjectName = sheetFileName.replace(".pdf", ".thumb.jpg");
                String path2 = minioUtil.uploadToMinio(thumbnailData, thumbObjectName,"image/jpeg");
                thumbs.put(i,minioUtil.getFileUrl(path2));
            }
        }
        result.put("pages",sheets);
        result.put("thumbnails",thumbs);
        return result;
    }

    /**
     * 根据文件扩展名创建相应的工作簿
     */
    private Workbook createWorkbook(InputStream is, String filename) throws IOException {
        if (filename.toLowerCase().endsWith(".xlsx")) {
            return new XSSFWorkbook(is);
        } else if (filename.toLowerCase().endsWith(".xls")) {
            return new HSSFWorkbook(is);
        } else {
            throw new IllegalArgumentException("不支持的文件格式: " + filename);
        }
    }

    private byte[] convertSheetToPdf(Sheet sheet) throws Exception {
        ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();

        // 👉 1. 加载中文字体（关键步骤）
        String fontPath = "fonts/nxcxs.ttf"; // 路径根据项目结构调整
        PdfFont font = PdfFontFactory.createFont(fontPath, PdfEncodings.IDENTITY_H); // 注意：用 IDENTITY_H 支持中文

        // 创建PDF文档
        PdfWriter writer = new PdfWriter(pdfOutputStream);
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document document = new Document(pdfDoc, PageSize.A4.rotate());
        document.setFont(font); // 👉 设置默认字体为中文字体

        // 设置边距
        document.setTopMargin(50);
        document.setBottomMargin(30);
        document.setLeftMargin(30);
        document.setRightMargin(30);

        // 添加标题（现在可以显示中文）
        Paragraph title = new Paragraph("工作表: " + sheet.getSheetName())
                .setFontSize(16)
                .setBold()
                .setTextAlignment(TextAlignment.CENTER)
                .setMarginBottom(20);
        // 不需要单独 setFont(font)，因为 document 已设置默认字体

        document.add(title);

        // 确定列数∂
        int maxColumns = 0;
        for (Row row : sheet) {
            if (row != null && row.getLastCellNum() > maxColumns) {
                maxColumns = row.getLastCellNum();
            }
        }

        if (maxColumns > 0) {
            Table table = new Table(UnitValue.createPercentArray(maxColumns));
            table.setWidth(UnitValue.createPercentValue(100));
            table.setMarginTop(10);

            for (Row row : sheet) {
                if (row == null) continue;

                for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    String cellValue = getCellValueAsString(cell);

                    Paragraph cellParagraph = new Paragraph(cellValue != null ? cellValue : "");
                    // cellParagraph.setFont(font); // 如果 document 没设字体，这里要设

                    com.itextpdf.layout.element.Cell pdfCell = new com.itextpdf.layout.element.Cell()
                            .add(cellParagraph)
                            .setPadding(3);

                    if (row.getRowNum() == 0) {
                        pdfCell.setBold(); // 加粗依赖字体支持
                    }

                    table.addCell(pdfCell);
                }
            }

            document.add(table);
        }

        document.close();
        return pdfOutputStream.toByteArray();
    }


    /**
     * 为指定 Sheet 生成缩略图（前几行 + 几列作为预览）
     */
    private byte[] generateThumbnail(Sheet sheet, int width, int height) throws IOException {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = image.createGraphics();

        // 启用抗锯齿
        g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

        // 背景白色
        g2d.setColor(Color.WHITE);
        g2d.fillRect(0, 0, width, height);

        // 字体设置，使用系统默认字体
        Font font = new Font("SansSerif", Font.PLAIN, 10);
        g2d.setFont(font);
        FontMetrics fm = g2d.getFontMetrics();

        // 单元格尺寸（动态计算）
        int margin = 5;
        int cellHeight = 20;
        int maxCols = Math.min(8, sheet.getRow(0) != null ? sheet.getRow(0).getLastCellNum() : 5); // 最多显示8列
        int maxRows = Math.min(15, sheet.getLastRowNum() + 1); // 最多显示15行

        double colWidthPx = (double)(width - 2 * margin) / maxCols;
        double rowHeightPx = (double)(height - 2 * margin) / maxRows;

        // 绘制表头和数据
        for (int r = 0; r < maxRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            for (int c = 0; c < maxCols; c++) {
                Cell cell = row.getCell(c);
                String value = getCellValueAsString(cell);

                double x = margin + c * colWidthPx;
                double y = margin + r * rowHeightPx + cellHeight - 4;

                // 绘制边框
                g2d.setColor(Color.LIGHT_GRAY);
                g2d.drawRect((int)x, (int)(y - cellHeight + 2), (int)colWidthPx, cellHeight);

                // 文本颜色
                g2d.setColor(Color.BLACK);
                // 居中裁剪文本
                String displayText = value.length() > 10 ? value.substring(0, 10) + ".." : value;
                int textWidth = fm.stringWidth(displayText);
                int textX = (int)(x + (colWidthPx - textWidth) / 2);
                g2d.drawString(displayText, Math.max((int)x + 3, textX), (int)y);
            }
        }

        g2d.dispose();

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "jpg", baos);
        return baos.toByteArray();
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return cell.toString();
        }
    }


    private String getBaseName(String filename) {
        int lastDot = filename.lastIndexOf('.');
        return lastDot > 0 ? filename.substring(0, lastDot) : filename;
    }
}
