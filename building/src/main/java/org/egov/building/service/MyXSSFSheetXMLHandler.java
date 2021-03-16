package org.egov.building.service;

import java.util.Arrays;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;

/**
 * A custom XSSFSheetXMLHandler that uses qName instead of localName. Also
 * handles numeric content so the format is removed when cell is created.
 */
public class MyXSSFSheetXMLHandler extends XSSFSheetXMLHandler {

    // Used to format numeric cell values.
    private short formatIndex;
    private String formatString;
    private boolean isCurrentTypeFormattedNumber = false;
    private String currentCellRef = null;
    private StringBuilder value = new StringBuilder(64);
    private final SheetContentsHandler output;
    /**
     * Table with the styles used for formatting
     */
    private Styles stylesTable;

    private final String[] numbericCellTypes = new String[] { "b", "e", "inlineStr", "s", "str" };

    public MyXSSFSheetXMLHandler(Styles styles, Comments comments, SharedStrings strings,
            SheetContentsHandler sheetContentsHandler, DataFormatter dataFormatter, boolean formulasNotResults) {
        super(styles, comments, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
        this.output = sheetContentsHandler;
        stylesTable = styles;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        super.startElement(null, qName, qName, attributes);
        if (this.isTextTag(qName)) {
            value.setLength(0);
            return;
        }
        currentCellRef = attributes.getValue("r");
        String cellType = attributes.getValue("t");
        String cellStyleStr = attributes.getValue("s");
        isCurrentTypeFormattedNumber = "c".equals(qName)
                && (cellType == null || !Arrays.stream(numbericCellTypes).anyMatch(cellType::equals));
        if (isCurrentTypeFormattedNumber) {
            XSSFCellStyle style = null;
            if (stylesTable != null) {
                if (cellStyleStr != null) {
                    int styleIndex = Integer.parseInt(cellStyleStr);
                    style = stylesTable.getStyleAt(styleIndex);
                } else if (stylesTable.getNumCellStyles() > 0) {
                    style = stylesTable.getStyleAt(0);
                }
            }
            if (style != null) {
                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();
                if (this.formatString == null)
                    this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
            }
            if (DateUtil.isADateFormat(formatIndex, formatString)) {
                isCurrentTypeFormattedNumber = false;
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        super.endElement(null, qName, qName);
        if (this.isTextTag(qName) && isCurrentTypeFormattedNumber) {
            this.output.cell(currentCellRef, value.toString(), null);
            isCurrentTypeFormattedNumber = false;
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        super.characters(ch, start, length);
        if (isCurrentTypeFormattedNumber) {
            value.append(ch, start, length);
        }
    }

    private boolean isTextTag(String name) {
        if ("v".equals(name)) {
            // Easy, normal v text tag
            return true;
        }
        if ("inlineStr".equals(name)) {
            // Easy inline string
            return true;
        }
        // It isn't a text tag
        return false;
    }
}