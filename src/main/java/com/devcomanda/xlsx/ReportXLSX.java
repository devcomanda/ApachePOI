package com.devcomanda.xlsx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ReportXLSX {

    private String filePath;

    public ReportXLSX(String filePath) {
        this.filePath = filePath;
    }

    public void create() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet report = workbook.createSheet("Report");
        Row title = report.createRow(0);

        Cell titleCell = createTitle(title, 1, "Рабочий стаж участников встречи");
        CellStyle titleStyle = createCellStyle(workbook, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        Font titleFont = createFont(workbook, (short)16, false, false, true);

        titleStyle.setFont(titleFont);
        titleCell.setCellStyle(titleStyle);

        report.addMergedRegion(new CellRangeAddress(0, 0, 1, 7));

        try (OutputStream os = new FileOutputStream(filePath)){
            workbook.write(os);
        }
    }

    private Cell createTitle(Row row, int column, String value) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);

        return cell;
    }

    private CellStyle createCellStyle(Workbook workbook, HorizontalAlignment hAlign, VerticalAlignment vAlign) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(hAlign);
        cellStyle.setVerticalAlignment(vAlign);

        return cellStyle;
    }

    private Font createFont(Workbook workbook, short height, boolean isItalic, boolean isUnderline, boolean isBold) {
        Font font = workbook.createFont();

        font.setFontHeightInPoints(height);
        font.setFontName("Calibri");
        font.setItalic(isItalic);
        font.setBold(isBold);
        font.setUnderline(isUnderline ? XSSFFont.U_SINGLE : XSSFFont.U_NONE);

        return font;
    }
}
