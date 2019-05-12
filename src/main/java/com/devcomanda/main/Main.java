package com.devcomanda.main;

import com.devcomanda.xlsx.ReportXLSX;

import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {
        ReportXLSX reportXLSX = new ReportXLSX("E:\\ApachePOITest\\myWorkbook.xlsx");
        reportXLSX.create();
    }
}
