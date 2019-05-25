package com.devcomanda.xlsx;

import com.devcomanda.user.User;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class ReportXLSX {

    private String filePath;

    public ReportXLSX(String filePath) {
        this.filePath = filePath;
    }

    public void create() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet report = workbook.createSheet("Report");
        Row title = report.createRow(0);

        Cell titleCell = createStringCell(title, 1, "Рабочий стаж участников встречи");
        CellStyle titleStyle = createCellStyle(workbook, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        Font titleFont = createFont(workbook, (short)16, false, false, true);

        titleStyle.setFont(titleFont);
        titleCell.setCellStyle(titleStyle);

        CellStyle baseStyle = workbook.createCellStyle();
        Font baseFont = createFont(workbook, (short)11, false, false, false);
        baseStyle.setFont(baseFont);
        baseStyle.setBorderTop(BorderStyle.THIN);
        baseStyle.setBorderRight(BorderStyle.THIN);
        baseStyle.setBorderBottom(BorderStyle.THIN);
        baseStyle.setBorderLeft(BorderStyle.THIN);

        CellStyle baseBoldStyle = workbook.createCellStyle();
        Font baseBoldFont = createFont(workbook, (short)11, false, false, true);
        baseBoldStyle.setFont(baseBoldFont);
        baseBoldStyle.setBorderTop(BorderStyle.THIN);
        baseBoldStyle.setBorderRight(BorderStyle.THIN);
        baseBoldStyle.setBorderBottom(BorderStyle.THIN);
        baseBoldStyle.setBorderLeft(BorderStyle.THIN);

        List<User> userList = generateUsers();

        Row rowName = report.createRow(2);
        Row rowExperience = report.createRow(3);

        createStringCell(rowName, 1, "Имя").setCellStyle(baseBoldStyle);
        createStringCell(rowExperience, 1, "Стаж").setCellStyle(baseBoldStyle);

        for (int index = 0; index < userList.size(); index++) {
            User user = userList.get(index);

            createStringCell(rowName, 2 + index, user.getName()).setCellStyle(baseStyle);
            createIntCell(rowExperience, 2 + index, user.getExperience()).setCellStyle(baseStyle);
        }

        rowExperience.createCell(3 + userList.size(), CellType.FORMULA).setCellFormula("SUM(C4:H4)");

        report.addMergedRegion(new CellRangeAddress(0, 0, 1, 7));

        try (OutputStream os = new FileOutputStream(filePath)){
            workbook.write(os);
        }
    }

    private List<User> generateUsers() {
        List<User> userList = new ArrayList<>();

        userList.add(new User("Денис", 63));
        userList.add(new User("Костя", 46));
        userList.add(new User("Ира", 54));
        userList.add(new User("Оля", 102));
        userList.add(new User("Кирилл", 96));
        userList.add(new User("Саша", 80));

        return userList;
    }

    private Cell createStringCell(Row row, int column, String value) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);

        return cell;
    }

    private Cell createIntCell(Row row, int column, Integer value) {
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
