package com.alibaba.excel.context;

import com.alibaba.excel.metadata.*;
import com.alibaba.excel.metadata.CellRange;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.metadata.TableStyle;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author jipengfei
 */
public class GenerateContextImpl implements GenerateContext {

    private Sheet currentSheet;

    private String currentSheetName;

    private ExcelTypeEnum excelType;

    private Workbook workbook;

    private OutputStream outputStream;

    private Map<Integer, Sheet> sheetMap = new ConcurrentHashMap<Integer, Sheet>();

    private Map<Integer, Table> tableMap = new ConcurrentHashMap<Integer, Table>();

    private CellStyle defaultCellStyle;

    private CellStyle currentHeadCellStyle;

    private CellStyle currentContentCellStyle;

    private ExcelHeadProperty excelHeadProperty;

    private boolean needHead = true;

    public GenerateContextImpl(InputStream templateInputStream, OutputStream out, ExcelTypeEnum excelType,
                               boolean needHead) throws IOException {
        if (ExcelTypeEnum.XLS.equals(excelType)) {
            if (templateInputStream == null) {
                this.workbook = new HSSFWorkbook();
            } else {
                this.workbook = new HSSFWorkbook(new POIFSFileSystem(templateInputStream));
            }
        } else {
            if (templateInputStream == null) {
                this.workbook = new SXSSFWorkbook(500);
            } else {
                this.workbook = new SXSSFWorkbook(new XSSFWorkbook(templateInputStream));
            }
        }
        this.outputStream = out;
        // 暂时不用 defaultCellStyle
        // this.defaultCellStyle = buildDefaultCellStyle();
        this.needHead = needHead;
    }

    private CellStyle buildDefaultHeadCellStyle() {
        CellStyle newCellStyle = buildDefaultCellStyle();
        Font font = this.workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        newCellStyle.setFont(font);
        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        newCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        return newCellStyle;
    }

    private CellStyle buildDefaultContentCellStyle() {
        CellStyle newCellStyle = buildDefaultCellStyle();
        Font font = this.workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 11);
        newCellStyle.setFont(font);
        return newCellStyle;
    }

    private CellStyle buildDefaultCellStyle() {
        CellStyle newCellStyle = this.workbook.createCellStyle();
        newCellStyle.setWrapText(true);
        newCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newCellStyle.setBorderBottom(BorderStyle.THIN);
        newCellStyle.setBorderRight(BorderStyle.THIN);
        return newCellStyle;
    }


    @Override
    public void buildCurrentSheet(com.alibaba.excel.metadata.Sheet sheet) {
        if (sheetMap.containsKey(sheet.getSheetNo())) {
            this.currentSheet = sheetMap.get(sheet.getSheetNo());
        } else {
            Sheet sheet1 = null;
            try {
                sheet1 = workbook.getSheetAt(sheet.getSheetNo());
            } catch (Exception e) {

            }
            if (sheet1 == null) {
                this.currentSheet = workbook.createSheet(
                        sheet.getSheetName() != null ? sheet.getSheetName() : sheet.getSheetNo() + "");
                this.currentSheet.setDefaultColumnWidth(20);
                for (Map.Entry<Integer, Integer> entry : sheet.getColumnWidthMap().entrySet()) {
                    currentSheet.setColumnWidth(entry.getKey(), entry.getValue());
                }
            } else {
                this.currentSheet = sheet1;
            }
            sheetMap.put(sheet.getSheetNo(), this.currentSheet);
            // 创建excelHeadProperty
            buildHead(sheet.getHead(), sheet.getClazz());
            buildTableStyle(sheet.getTableStyle());
            if (needHead && excelHeadProperty != null) {
                appendHeadToExcel();
            }
        }

    }

    private void buildHead(List<List<String>> head, Class<? extends BaseRowModel> clazz) {
        if (head != null || clazz != null) {
            excelHeadProperty = new ExcelHeadProperty(clazz, head);
        }
    }

    public void appendHeadToExcel() {
        if (this.excelHeadProperty.getHead() != null && this.excelHeadProperty.getHead().size() > 0) {
            List<CellRange> list = this.excelHeadProperty.getCellRangeModels();
            int n = currentSheet.getLastRowNum();
            if (n > 0) {
                n = n + 4;
            }
            for (CellRange cellRangeModel : list) {
                CellRangeAddress cra = new CellRangeAddress(cellRangeModel.getFirstRow() + n,
                        cellRangeModel.getLastRow() + n,
                        cellRangeModel.getFirstCol(), cellRangeModel.getLastCol());
                currentSheet.addMergedRegion(cra);
            }
            int i = n;
            for (; i < this.excelHeadProperty.getRowNum() + n; i++) {
                Row row = currentSheet.createRow(i);
                addOneRowOfHeadDataToExcel(row, this.excelHeadProperty.getHeadByRowNum(i - n));
            }
        }
    }

    private void addOneRowOfHeadDataToExcel(Row row, List<String> headByRowNum) {
        if (headByRowNum != null && headByRowNum.size() > 0) {
            for (int i = 0; i < headByRowNum.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(this.getCurrentHeadCellStyle());
                cell.setCellValue(headByRowNum.get(i));
            }
        }
    }

    /**
     * 只有不为空时能赋值默认值
     *
     * @param tableStyle
     */
    private void buildTableStyle(TableStyle tableStyle) {
        if (tableStyle != null) {
            CellStyle headStyle = tableStyle.getTableHeadCellStyle();
            if (headStyle == null) {
                headStyle = buildDefaultCellStyle();
                if (tableStyle.getTableHeadFont() != null) {
                    Font font = this.workbook.createFont();
                    font.setFontName(tableStyle.getTableHeadFont().getFontName());
                    font.setFontHeightInPoints(tableStyle.getTableHeadFont().getFontHeightInPoints());
                    font.setBold(tableStyle.getTableHeadFont().isBold());
                    headStyle.setFont(font);
                }
                if (tableStyle.getTableHeadBackGroundColor() != null) {
                    headStyle.setFillForegroundColor(tableStyle.getTableHeadBackGroundColor().getIndex());
                }
            }
            this.currentHeadCellStyle = headStyle;

            CellStyle contentStyle = tableStyle.getTableContentCellStyle();
            if (contentStyle == null) {
                contentStyle = buildDefaultCellStyle();
                if (tableStyle.getTableContentFont() != null) {
                    Font font = this.workbook.createFont();
                    font.setFontName(tableStyle.getTableContentFont().getFontName());
                    font.setFontHeightInPoints(tableStyle.getTableContentFont().getFontHeightInPoints());
                    font.setBold(tableStyle.getTableContentFont().isBold());
                    contentStyle.setFont(font);
                }
                if (tableStyle.getTableContentBackGroundColor() != null) {
                    contentStyle.setFillForegroundColor(tableStyle.getTableContentBackGroundColor().getIndex());
                }
            }
            this.currentContentCellStyle = contentStyle;

        } else {
            this.currentHeadCellStyle = buildDefaultHeadCellStyle();
            this.currentContentCellStyle = buildDefaultContentCellStyle();
        }
    }

    @Override
    public void buildTable(Table table) {
        if (!tableMap.containsKey(table.getTableNo())) {
            buildHead(table.getHead(), table.getClazz());
            tableMap.put(table.getTableNo(), table);
            buildTableStyle(table.getTableStyle());
            if (needHead && excelHeadProperty != null) {
                appendHeadToExcel();
            }
        }

    }

    @Override
    public void buildMergeCells(List<CellRange> mergeCells) {
        if (mergeCells != null && mergeCells.size() > 0) {
            for (CellRange cellRange : mergeCells) {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(cellRange.getFirstRow(),
                        cellRange.getLastRow(),
                        cellRange.getFirstCol(),
                        cellRange.getLastCol());
                currentSheet.addMergedRegion(cellRangeAddress);
            }
        }
    }

    @Override
    public ExcelHeadProperty getExcelHeadProperty() {
        return this.excelHeadProperty;
    }

    @Override
    public boolean needHead() {
        return this.needHead;
    }

    @Override
    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public String getCurrentSheetName() {
        return currentSheetName;
    }

    public void setCurrentSheetName(String currentSheetName) {
        this.currentSheetName = currentSheetName;
    }

    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    @Override
    public OutputStream getOutputStream() {
        return outputStream;
    }

    @Override
    public CellStyle getCurrentHeadCellStyle() {
        return this.currentHeadCellStyle == null ? defaultCellStyle : this.currentHeadCellStyle;
    }

    @Override
    public CellStyle getCurrentContentStyle() {
        return this.currentContentCellStyle;
    }

    @Override
    public Workbook getWorkbook() {
        return workbook;
    }

}
