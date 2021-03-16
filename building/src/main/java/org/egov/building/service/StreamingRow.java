package org.egov.building.service;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;

public class StreamingRow extends SXSSFRow {

    private int rowNum;

    public StreamingRow() {
        super(null);
    }

    @Override
    public int getRowNum() {
        return rowNum;
    }

    @Override
    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    @Override
    public SXSSFCell getCell(int cellnum) {
        return getCell(cellnum, MissingCellPolicy.RETURN_BLANK_AS_NULL);
    }
}
