package org.egov.building.service;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * Uses the XSSF Event SAX helpers to do most of the work of parsing the Sheet
 * XML, and outputs the contents as a (basic) CSV.
 */
public class StreamingSheetContentsHandler implements SheetContentsHandler {

    private StreamingRowProcessor mProcessor;

    /**
     * @param readExcelStreamingService
     */
    StreamingSheetContentsHandler(StreamingRowProcessor processor) {
        this.mProcessor = processor;
    }

    private int currentRow = -1;
    private int currentCol = -1;
    private StreamingRow currentStreamingRow;

    @Override
    public void startRow(int rowNum) {
        // Prepare for this row
        currentRow = rowNum;
        currentCol = -1;
        this.currentStreamingRow = new StreamingRow();
        this.currentStreamingRow.setRowNum(rowNum);
    }

    @Override
    public void endRow(int rowNum) {
        this.mProcessor.processRow(this.currentStreamingRow);
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {

        // gracefully handle missing CellRef here in a similar way as XSSFCell does
        if (cellReference == null) {
            cellReference = new CellAddress(currentRow, currentCol).formatAsString();
        }

        // Did we miss any cells?
        int thisCol = (new CellReference(cellReference)).getCol();
        for (int i = currentCol + 1; i > -1 && i < thisCol; i++) {
            this.currentStreamingRow.createCell(i);
        }
        // this.currentStreamingRow.setCellData(thisCol, formattedValue);
        SXSSFCell cell = this.currentStreamingRow.createCell(thisCol);
        try {
            cell.setCellValue(Double.parseDouble(formattedValue));
            cell.setCellType(CellType.NUMERIC);
        } catch (Exception e) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(formattedValue);
        }
        currentCol = thisCol;
    }

    public interface StreamingRowProcessor {
        public void processRow(Row row);
    }
}