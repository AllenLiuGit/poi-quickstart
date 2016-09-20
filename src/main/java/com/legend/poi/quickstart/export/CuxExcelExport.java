package com.legend.poi.quickstart.export;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by allen on 9/20/16.
 */
public class CuxExcelExport {
    public static final int CFG_MAX_SHEET_SIZE = 65535; // Max size per sheet
    public static final String CFG_SHEET_PREFIX = "Sheet_"; // Sheet name prefix

    private int currentRowNum; // Current data size, used to calculate new sheet count
    private List<String> columnTitleList; // Column title

    private OutputStream outputStream; // Output stream
    private SXSSFWorkbook workbook; // Workbook
    private List<Sheet> sheetList; // All sheets, exceeds 65535, split into another sheet
    private Sheet currentSheet; // Current sheet

    /**
     * Constructor
     *
     * @param outputStream
     * @param rowAccessWindowSize
     * @param columnTitleList Tile of all rows
     */
    public CuxExcelExport(OutputStream outputStream, int rowAccessWindowSize, List<String> columnTitleList) {
        // Validate
        if (outputStream == null || rowAccessWindowSize < 0) {
            throw new IllegalArgumentException("Please provide valid output stream and row limit to flush temporary.");
        }
        if (columnTitleList == null || columnTitleList.size() == 0) {
            throw new IllegalArgumentException("Please provide valid column title list.");
        }

        // Initialize
        this.currentRowNum = 0;
        this.columnTitleList = columnTitleList;

        this.outputStream = outputStream;
        this.workbook = new SXSSFWorkbook(rowAccessWindowSize);
        this.sheetList = new ArrayList<Sheet>();

    }

    /**
     * 添加行数据
     * @param rowDataList
     */
    private void addRowDataInternal(List<String> rowDataList) {
        // Validate
        if (rowDataList == null || columnTitleList == null || rowDataList.size() != columnTitleList.size()) {
            throw new RuntimeException("Please review column title information and row data, their size are not match.");
        }

        // Create row and set cell values
        int rowIndex = this.currentRowNum%CFG_MAX_SHEET_SIZE;
        this.currentRowNum++;
        Row row = this.currentSheet.createRow(rowIndex);

        for (int i = 0; i < rowDataList.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(rowDataList.get(i)); // Set cell value
        }
    }

    /**
     * Create title
     */
    private void addColumnTitleInternal() {
        if (this.columnTitleList == null || this.columnTitleList.size() == 0) {
            throw new RuntimeException("Please initialize title list first.");
        }

        // Create row and set cell values
        if (this.currentSheet == null || (this.currentRowNum % CFG_MAX_SHEET_SIZE) == 0) {
            this.currentSheet = this.workbook.createSheet(CFG_SHEET_PREFIX + this.sheetList.size()); // Create new sheet
            this.sheetList.add(this.currentSheet); // Add to sheet list
            this.addRowDataInternal(this.columnTitleList); // Add column title
        }

    }

    /**
     * Add row data
     *
     * @param rowDataList
     */
    public void addRowData(List<List<String>> rowDataList) {
        if (rowDataList == null || rowDataList.size() == 0) {
            throw new IllegalArgumentException("Please provide valid row data list.");
        }

        for (List<String> rowData : rowDataList) {
            this.addColumnTitleInternal(); // Add column title or not
            this.addRowDataInternal(rowData); // Add row data
        }
    }

    /**
     * Flush and close
     */
    public void flushThenClose() throws IOException {
        // Write to output stream
        this.workbook.write(this.outputStream);

        // Dispose of temporary files backing this workbook on disk
        this.workbook.dispose();
    }

}
