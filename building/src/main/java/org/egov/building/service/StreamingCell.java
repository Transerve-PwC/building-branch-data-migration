package org.egov.building.service;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;

public class StreamingCell extends SXSSFCell {

    public StreamingCell(SXSSFRow row, CellType cellType) {
        super(row, cellType);
    }
}
