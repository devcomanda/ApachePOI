package com.devcomanda.main;

import com.devcomanda.docx.ReportDOCX;
import com.devcomanda.xlsx.ReportXLSX;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        /*ReportXLSX reportXLSX = new ReportXLSX("E:\\ApachePOITest\\myWorkbook.xlsx");
        reportXLSX.create();*/

        ReportDOCX reportDOCX = new ReportDOCX("E:\\ApachePOITest\\recipe.docx");
        reportDOCX.create();
    }
}
