package com.legend.poi.quickstart.demo;

import com.legend.poi.quickstart.export.CuxExcelExport;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by allen on 9/20/16.
 */
public class CuxExcelExportExample {
    public static void main(String[] args) throws IOException{
        FileOutputStream out = new FileOutputStream("/Users/allen/Downloads/CuxExcelExport.xlsx");
        List<String> columnTitleList = new ArrayList<String>();
        columnTitleList.add("姓名");
        columnTitleList.add("年龄");

        CuxExcelExport cuxExcelExport = new CuxExcelExport(out, 500, columnTitleList);
        List<List<String>> rowDataList = new ArrayList<List<String>>();
        for (int i = 0; i < 65535; i++) {
            List<String> rowData = new ArrayList<String>();
            for (int j = 0; j < 2; j++) {
                rowData.add("Cell_" + i + ", " + j);
            }
            rowDataList.add(rowData);
        }

        cuxExcelExport.addRowData(rowDataList);
        cuxExcelExport.addRowData(rowDataList);
        cuxExcelExport.addRowData(rowDataList);
        cuxExcelExport.addRowData(rowDataList);
        cuxExcelExport.addRowData(rowDataList);

        cuxExcelExport.flushThenClose();
    }
}
