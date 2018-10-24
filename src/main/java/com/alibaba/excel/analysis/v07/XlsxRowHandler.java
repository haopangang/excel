package com.alibaba.excel.analysis.v07;

import com.alibaba.excel.annotation.FieldType;
import com.alibaba.excel.constant.ExcelXmlConstants;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventRegisterCenter;
import com.alibaba.excel.event.OneRowAnalysisFinishEvent;
import com.alibaba.excel.util.PositionUtils;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Arrays;

import static com.alibaba.excel.constant.ExcelXmlConstants.*;

/**
 * 解析之后对各个
 * @author jipengfei
 */
public class XlsxRowHandler extends DefaultHandler {

    private String currentCellIndex;

    private FieldType currentCellType;

    private int curRow;

    private int curCol;

    /**
     * 当前Excel中一行的所有数据
     */
    private String[] curRowContent = new String[20];

    private String currentCellValue;

    private SharedStringsTable sst;

    private AnalysisContext analysisContext;

    private AnalysisEventRegisterCenter registerCenter;

    public XlsxRowHandler(AnalysisEventRegisterCenter registerCenter, SharedStringsTable sst,
                          AnalysisContext analysisContext) {
        this.registerCenter = registerCenter;
        this.analysisContext = analysisContext;
        this.sst = sst;

    }

    /**
     * 开始解析元素时，触发的事件
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     * @throws SAXException
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

        setTotalRowCount(name, attributes);

        startCell(name, attributes);

        startCellValue(name);

    }

    private void startCellValue(String name) {
        if (name.equals(CELL_VALUE_TAG) || name.equals(CELL_VALUE_TAG_1)) {
            // initialize current cell value
            currentCellValue = "";
        }
    }

    /**
     * 开始一个cell的值的获取
     * @param name
     * @param attributes
     */
    private void startCell(String name, Attributes attributes) {
        if (ExcelXmlConstants.CELL_TAG.equals(name)) {
            currentCellIndex = attributes.getValue(ExcelXmlConstants.POSITION);
            int nextRow = PositionUtils.getRow(currentCellIndex);
            if (nextRow > curRow) {
                curRow = nextRow;
                // endRow(ROW_TAG);
            }
            analysisContext.setCurrentRowNum(curRow);
            curCol = PositionUtils.getCol(currentCellIndex);

            String cellType = attributes.getValue("t");
            currentCellType = FieldType.EMPTY;
            if (cellType != null && cellType.equals("s")) {
                currentCellType = FieldType.STRING;
            }
            //if ("6".equals(attributes.getValue("s"))) {
            //    // date
            //    currentCellType = FieldType.DATE;
            //}

        }
    }

    /**
     * 结束一个cellVal执行的方法
     * @param name
     * @throws SAXException
     */
    private void endCellValue(String name) throws SAXException {
        // ensure size
        if (curCol >= curRowContent.length) {
            curRowContent = Arrays.copyOf(curRowContent, (int)(curCol * 1.5));
        }
        if (CELL_VALUE_TAG.equals(name)) {

            switch (currentCellType) {
                case STRING:
                    int idx = Integer.parseInt(currentCellValue);
                    currentCellValue = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    currentCellType = FieldType.EMPTY;
                    break;
                //case DATE:
                //    Date dateVal = HSSFDateUtil.getJavaDate(Double.parseDouble(currentCellValue),
                //        analysisContext.use1904WindowDate());
                //    currentCellValue = TypeUtil.getDefaultDateString(dateVal);
                //    currentCellType = FieldType.EMPTY;
                //    break;
            }
            curRowContent[curCol] = currentCellValue;
        } else if (CELL_VALUE_TAG_1.equals(name)) {
            curRowContent[curCol] = currentCellValue;
        }
    }

    /**
     * 一个元素结束之后，执行的方法
     * @param uri
     * @param localName
     * @param name
     * @throws SAXException
     */
    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {

        endRow(name);
        endCellValue(name);
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {

        currentCellValue += new String(ch, start, length);

    }


    /**
     * 根据POI方法解析出来的Excel，获取其总共有多上行
     *
     * @param name
     * @param attributes
     */
    private void setTotalRowCount(String name, Attributes attributes) {
        if (DIMENSION.equals(name)) {
            String d = attributes.getValue(DIMENSION_REF);
            String totalStr = d.substring(d.indexOf(":") + 1, d.length());
            String c = totalStr.toUpperCase().replaceAll("[A-Z]", "");
            analysisContext.setTotalCount(Integer.parseInt(c));
        }

    }

    /**
     * 如果是一行的数据结束之后，触发我们的结束方法
     *
     * @param name
     */
    private void endRow(String name) {

        if (name.equals(ROW_TAG)) {
            registerCenter.notifyListeners(new OneRowAnalysisFinishEvent(Arrays.asList(curRowContent)));
            curRowContent = new String[20];
        }
    }

}

