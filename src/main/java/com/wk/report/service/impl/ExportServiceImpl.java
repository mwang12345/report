package com.wk.report.service.impl;

import com.wk.report.pojo.HeaderValue;
import com.wk.report.pojo.PrintRow;
import com.wk.report.pojo.PrintTD;
import com.wk.report.pojo.ReportIn;
import com.wk.report.service.ExportService;
import com.wk.report.service.ReportService;
import freemarker.template.Configuration;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

@Service
public class ExportServiceImpl implements ExportService {

    private static int PAGE_SIZE = 30;

    private Configuration cfg;

    private static Logger logger = LoggerFactory.getLogger(ExportServiceImpl.class);

    @Value("${zdxf.report.tpl}")
    private String tplPath;

    @Value("${zdxf.report.template}")
    private String templatePath;

    @Value("${zdxf.report.genPath}")
    private String genPath;

    @Autowired
    private ReportService reportService;

    @PostConstruct
    public void init() {

        cfg = new Configuration(Configuration.DEFAULT_INCOMPATIBLE_IMPROVEMENTS);
        try {
            cfg.setDirectoryForTemplateLoading(new File(tplPath));
        } catch (IOException e) {
            logger.error(e.getMessage());
        }
    }

    /**
     * 打印报表转成html内容
     * @return
     * @throws Exception
     */
    public HSSFWorkbook export(ReportIn reportIn, String fileName) throws Exception {
        String templateFilePath = templatePath + "//" + reportIn.getFileName() + ".xls";
//        String targetFilePath = genPath + "//" + fileName;
//        FileUtils.copyFile(templateFilePath, targetFilePath);
        HSSFWorkbook oldWb = new HSSFWorkbook(new FileInputStream(templateFilePath));
        HSSFWorkbook newWb = new HSSFWorkbook();
        exportData(oldWb, newWb, reportIn);
        return newWb;
    }

    /**
     * 将Excel模型对象转换成表格封装对象
     * @return
     */
    private void exportData(HSSFWorkbook oldWb, HSSFWorkbook newWb, ReportIn reportIn) throws Exception{

        if (newWb == null) {
            return;
        }
        HSSFSheet destSheet = newWb.createSheet();
        HSSFSheet srcSheet = oldWb.getSheetAt(0);

        // 复制表头
        copySheet(oldWb, newWb, reportIn, srcSheet, destSheet, reportIn.getShowIndexs());

        int firstRowNum = srcSheet.getFirstRowNum();
        int lastRowNum = srcSheet.getLastRowNum();
        boolean notFillData = true;
        int startFillDataIdx = 0;
        String dataSampleIndex = "";
        for(int rowIdx = firstRowNum; rowIdx < lastRowNum; rowIdx++) {
            Row oldRow = srcSheet.getRow(rowIdx);
            if (dataSampleIndex.indexOf(",") == -1) {
                dataSampleIndex = getDataSampleIndex(oldRow);
            }
            if (notFillData && dataSampleIndex.indexOf(",") != -1) {
                startFillDataIdx = rowIdx - 1;
                notFillData = false;
            }
        }
        if (reportIn.getHeaderValues() != null) {
            startFillDataIdx = startFillDataIdx + reportIn.getHeaderValues().length;
        }
        createValueRows(newWb, destSheet, startFillDataIdx, reportIn, dataSampleIndex);
    }

    /**
     * 打印报表转成html内容
     * @return
     * @throws Exception
     */
    public HSSFWorkbook exportWithoutTemplate(ReportIn reportIn) throws Exception {
        HSSFWorkbook newWb = new HSSFWorkbook();
        exportData(newWb, reportIn);
        return newWb;
    }

    /**
     * 将Excel模型对象转换成表格封装对象
     * @return
     */
    private void exportData(HSSFWorkbook newWb, ReportIn reportIn) throws Exception{

        if (newWb == null) {
            return;
        }
        HSSFSheet destSheet = newWb.createSheet();

        createValueRows(newWb, destSheet, 0, reportIn, "");
    }

    private List<CellRangeAddress> getSortRange(HSSFSheet sheet) {
        List<CellRangeAddress> list = sheet.getMergedRegions();
        if (list == null) {
            return null;
        }
        list.sort(new Comparator<CellRangeAddress>() {
            @Override
            public int compare(CellRangeAddress o1, CellRangeAddress o2) {
                return o1.getFirstRow() - o2.getFirstRow();
            }
        });
        return list;
    }

    /**
     * 插入自定义表头
     * @param wb
     * @param sheet
     * @param srcSheet
     * @param beginRowIndex
     * @param reportIn
     * @return
     */
    private int createCustomerHeaders(HSSFWorkbook wb, HSSFSheet sheet, HSSFSheet srcSheet, int beginRowIndex, ReportIn reportIn) {

        if (reportIn.getHeaderValues() == null) {
            return 0;
        }
        int mergeFirstCol = -1;
        int mergeLastCol = -1;
        List<CellRangeAddress> ranges = getSortRange(srcSheet);
        if (ranges != null) {
            mergeFirstCol = ranges.get(0).getFirstColumn();
            mergeLastCol = ranges.get(0).getLastColumn();
        } else {
            mergeFirstCol = 1;
            mergeLastCol = reportIn.getValues()[0].split(";").length + 1;
        }
        //样式
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //背景色白色
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.WHITE.index);
        //边框
        style.setBorderTop(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        HSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short)10);
        style.setFont(font);

        List<Integer> newRowNums = new ArrayList<>();
        for (int index = 0; index < reportIn.getHeaderValues().length; index++){
            Row newRow = sheet.createRow(beginRowIndex + index);
            newRow.setHeightInPoints(20);
            for (int colIdx = mergeFirstCol; colIdx <= mergeLastCol; colIdx++) {
                Cell cellLabel = newRow.createCell(colIdx);
                cellLabel.setCellStyle(style);
                if (colIdx == mergeFirstCol) {
                    HeaderValue headerValue = reportIn.getHeaderValues()[index];
                    cellLabel.setCellValue(headerValue.getLabel() + headerValue.getValue());
                    newRowNums.add(newRow.getRowNum());
                } else {
                    cellLabel.setCellValue("");
                }
            }
        }
        for (Integer rowNum : newRowNums) {
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, mergeFirstCol, mergeLastCol));
        }
        return reportIn.getHeaderValues().length;
    }

    /**
     * 表单复制
     * @param srcWb
     * @param wb
     * @param reportIn
     * @param srcSheet
     * @param destSheet
     * @param showIndexs
     */
    private void copySheet(HSSFWorkbook srcWb, HSSFWorkbook wb, ReportIn reportIn, HSSFSheet srcSheet, HSSFSheet destSheet, Integer[] showIndexs) {
        Iterator<Row> rowIt = srcSheet.rowIterator();

        List<CellRangeAddress> ranges = getSortRange(srcSheet);
        int customerHeaderLength = 0;
        if (reportIn.getHeaderValues() != null) {
            customerHeaderLength = reportIn.getHeaderValues().length;
        }
        int startFirstHeader = 0;
        boolean startCustomer = false;
        while(rowIt.hasNext()) {
            Row srcRow = rowIt.next();
            int rowNum = srcRow.getRowNum();
            if (startCustomer) {
                rowNum = rowNum + customerHeaderLength;
            }
            // 处理自定义表头
            if (startFirstHeader == 1) {
                startCustomer = true;
                startFirstHeader = 2;
                this.createCustomerHeaders(wb, destSheet, srcSheet, rowNum, reportIn);
                rowNum = rowNum + customerHeaderLength;
            }
            Row destRow = destSheet.createRow(rowNum);
            destRow.setHeight(srcRow.getHeight());
            Iterator<Cell> srcCellIt = srcRow.cellIterator();
            if (srcRow.getRowNum() > reportIn.getHeaderRows()) {
                if (ranges != null) {
                    // 处理合并单元格
                    for (int index = 0; index < ranges.size(); index++) {
                        CellRangeAddress oldRange = ranges.get(index);
                        CellRangeAddress range = new CellRangeAddress(oldRange.getFirstRow(), oldRange.getLastRow(), oldRange.getFirstColumn(), oldRange.getLastColumn());
                        if (index != 0) {
                            range.setFirstRow(range.getFirstRow() + customerHeaderLength);
                            range.setLastRow(range.getLastRow() + customerHeaderLength);
                        }
                        if (!destSheet.getMergedRegions().contains(range)) {
                            try{
                                destSheet.addMergedRegion(range);
                            }catch (Exception e) {
                                logger.error(e.getMessage());
                            }
                        }
                    }
                }
                return;
            }
            int beginCellIdx = -1;
            while(srcCellIt.hasNext()) {
                if (startFirstHeader == 0) {
                    startFirstHeader = 1;
                }
                Cell srcCell = srcCellIt.next();
                if (beginCellIdx == -1) {
                    beginCellIdx = srcCell.getColumnIndex();
                }
                int headerShowIdx = srcCell.getColumnIndex() - beginCellIdx;
                if (!ReportServiceImpl.showTD(headerShowIdx, showIndexs)) {
                    continue;
                }
                // 复制样式
                // 绘制标题单元格
                HSSFCellStyle srcCellStyle = (HSSFCellStyle)srcCell.getCellStyle();
                Cell destCell = destRow.createCell(srcCell.getColumnIndex());
                String value = srcCell.toString();
                destCell.setCellValue(value);
                String mergeRange = getMergeRange(ranges, destCell, customerHeaderLength);
                handleSlash(wb, destSheet, destCell, value, Integer.parseInt(mergeRange.split(",")[0]), Integer.parseInt(mergeRange.split(",")[1]));
                destCell.setCellStyle(getStyle(srcWb, wb, srcCellStyle));
            }
        }
    }

    /**
     * 解析Excel合并单元格，获取跨行/列数量，供画斜线使用
     * @param ranges
     * @param destCell
     * @return
     */
    private String getMergeRange(List<CellRangeAddress> ranges, Cell destCell, int customerHeaderRows) {
        if (ranges == null || destCell == null) {
            return null;
        }
        int rowIdx = destCell.getRowIndex();
        int colIdx = destCell.getColumnIndex();
        int rowRange = 0;
        int colRange = 0;
        for (CellRangeAddress range : ranges) {
            if (rowIdx >= range.getFirstRow() + customerHeaderRows
                    && rowIdx <= range.getLastRow() + customerHeaderRows
                    && colIdx >= range.getFirstColumn()
                    && colIdx <= range.getLastColumn()) {
                rowRange = range.getLastRow() - range.getFirstRow();
                colRange = range.getLastColumn() - range.getFirstColumn();
                return rowRange + "," + colRange;
            }
        }
        return "0,0";
    }

    /**
     * 斜线字符串拼接
     * @param value
     * @return
     */
    private String appendSlashValue(String value, boolean left) {
        if (StringUtils.isEmpty(value)) {
            return value;
        }
        int appendLength = 4 - value.length();
        String tmp = "";
        for (int index = 0; index < appendLength; index++) {
            tmp = tmp + " ";
        }
        if (left) {
            return value + tmp;
        } else {
            return tmp + value;
        }

    }

    /**
     * 处理表格斜线
     */
    private void handleSlash(HSSFWorkbook wb, HSSFSheet sheet, Cell cell, String value, int mergeRows, int mergeCols) {
        // 处理表头斜线
        if(value.indexOf("\\\\") == -1) {
            return;
        }
        String value1 = value.split("\\\\")[0];
        String value2 = value.split("\\\\")[2];

        value1 = value1 + "   ";
        value2 = "   " + value2;
        String cellData = value1 + value2;
        int valueLength = cellData.length() ;
        cell.setCellValue(cellData);
        CreationHelper helper = wb.getCreationHelper();
        HSSFPatriarch drawing = (HSSFPatriarch) sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        // 设置斜线的开始位置
        int colIdx = cell.getColumnIndex();
        int rowIdx = cell.getRowIndex();
        anchor.setCol1(colIdx);
        anchor.setRow1(rowIdx);
        // 设置斜线的结束位置
        anchor.setCol2(colIdx + 1 + mergeCols);
        anchor.setRow2(rowIdx + 1 + mergeRows);
        HSSFSimpleShape shape = drawing.createSimpleShape((HSSFClientAnchor) anchor);
        // 设置线宽
        shape.setLineWidth(10);
        // 设置线的颜色
        shape.setLineStyleColor(0, 0, 0);
        shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
        shape.setLineStyle(HSSFSimpleShape.LINESTYLE_SOLID);

        int width = (valueLength - 1) * 3 * 256;
        sheet.setColumnWidth(colIdx, width);
    }

    /**
     * 获取NA下标
     * @param row
     * @return
     */
    private String getDataSampleIndex(Row row) {
        if (row == null) {
            return null;
        }
        Iterator<Cell> cellIt = row.cellIterator();
        StringJoiner stringJoiner = new StringJoiner(",");
        while(cellIt.hasNext()) {
            Cell cell = cellIt.next();
            if ("na".equals(cell.toString().toLowerCase())) {
                stringJoiner.add(cell.getColumnIndex() + "");
            }
        }
        return stringJoiner.toString();
    }

    /**
     * 导出Value业务数据
     * @param wb
     * @param sheet
     * @param curRowNum
     * @param reportIn
     * @param dataSampleIndex
     * @return
     */
    private int createValueRows(HSSFWorkbook wb, HSSFSheet sheet, int curRowNum, ReportIn reportIn, String dataSampleIndex) {
        String[] values = reportIn.getValues();
        if (values == null || values.length == 0) {
            return 0;
        }
        sheet.setDefaultColumnWidth(20);
        String[] sampleIndex = null;
        if (dataSampleIndex.indexOf(",") != -1) {
            sampleIndex = dataSampleIndex.split(",");
        }
        // 开始填充数据
        HSSFCellStyle style = getStyle(wb, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, BorderStyle.THIN,"宋体",(short)14, IndexedColors.AUTOMATIC.getIndex(), false);
        // 合并行单元格下标数据
        List<String> mergeColIndexs = new ArrayList<>();
        List<String> mergeRowIndexs = new ArrayList<>();
        for(int rowIndex = 0; rowIndex < values.length; rowIndex++) {
            String[] valueItem = values[rowIndex].split(";");
            Row newRow = sheet.createRow(curRowNum + rowIndex + 1);
            newRow.setHeightInPoints(20);
            int newRowNum = newRow.getRowNum();
            Integer[] showIndex = reportIn.getShowIndexs();
            int cellColIdx = 0;
            int mergeColStart = Integer.MAX_VALUE;
            int mergeColEnd = -1;
            for (int colIndex = 0; colIndex < valueItem.length; colIndex++){
                if (sampleIndex != null) {
                    cellColIdx = Integer.parseInt(sampleIndex[colIndex]);
                } else {
                    cellColIdx = colIndex + 1;
                }
                if (!ReportServiceImpl.showTD(colIndex, showIndex)) {
                    continue;
                }
                Cell newCell = newRow.createCell(cellColIdx);
                String value = valueItem[colIndex];
                int cpIdx = value.indexOf("cp=");
                int mergeRows = 0;
                int mergeCols = 0;
                if (cpIdx != -1) {
                    int mergeCol = Integer.parseInt(value.substring(cpIdx + 3));
                    mergeColStart = cellColIdx;
                    mergeColEnd = mergeColStart + mergeCol - 1;
                    mergeCols = mergeColEnd - mergeColStart + 1;
                    mergeColIndexs.add(newRowNum + "," + newRowNum + "," + mergeColStart + "," + mergeColEnd);
                    value = value.substring(0, cpIdx - 1);
                }
                int rpIdx = value.indexOf("rp=");
                if (rpIdx != -1) {
                    int mergeRow = Integer.parseInt(value.substring(rpIdx + 3));
                    int mergeRowStart = newRowNum;
                    int mergeRowEnd = newRowNum + mergeRow - 1;
                    mergeRows = mergeRowEnd - mergeRowEnd + 1;
                    mergeRowIndexs.add(mergeRowStart + "," + mergeRowEnd + "," + cellColIdx + "," + cellColIdx);
                    value = value.substring(0, rpIdx - 1);
                }
                try{
                    newCell.setCellValue(Double.parseDouble(value));
                }catch(Exception e) {
                    newCell.setCellValue(value);
                }
                newCell.setCellStyle(style);

                // 添加斜线
                handleSlash(wb, sheet, newCell, value, mergeRows, mergeCols);
            }
        }
        // 合并列单元格
        for (String mergeIndex : mergeColIndexs) {
            handleMergeCell(mergeIndex, sheet, style);
        }
        // 合并行单元格
        for (String mergeIndex : mergeRowIndexs) {
            handleMergeCell(mergeIndex, sheet, style);
        }
        return values.length;
    }

    /**
     * 合并单元格
     * @param mergeIndex
     * @param sheet
     * @param style
     */
    private void handleMergeCell(String mergeIndex, HSSFSheet sheet, HSSFCellStyle style) {
        try{
            String[] mergeArray = mergeIndex.split(",");
            int rowStart = Integer.parseInt(mergeArray[0]);
            int rowEnd = Integer.parseInt(mergeArray[1]);
            int colStart = Integer.parseInt(mergeArray[2]);
            int colEnd = Integer.parseInt(mergeArray[3]);
            sheet.addMergedRegion(new CellRangeAddress(rowStart,rowEnd,colStart,colEnd));

            // 设置样式，主要解决合并后的边框丢失
            for (int rowIdx = rowStart; rowIdx <= rowEnd; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                for (int colIdx = colStart; colIdx <= colEnd; colIdx++) {
                    Cell cell = row.getCell(colIdx);
                    if (cell == null) {
                        cell = row.createCell(colIdx);
                        cell.setCellValue(" ");
                    }
                    cell.setCellStyle(style);
                }
            }
        }catch (Exception e) {
            e.printStackTrace();
            logger.error(e.getMessage());
        }
    }

    /**
     * 获取样式
     * @param workbook
     * @param horizontalAlignment
     * @param verticalAlignment
     * @param borderStyle
     * @param fontName
     * @param fontHeight
     * @param fontColor
     * @param bolder
     * @return
     */
    private static HSSFCellStyle getStyle(HSSFWorkbook workbook, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment, BorderStyle borderStyle, String fontName, short fontHeight, short fontColor, Boolean bolder) {
        // TODO Auto-generated method stub
        HSSFCellStyle style = workbook.createCellStyle();

        style.setAlignment(horizontalAlignment);
        style.setVerticalAlignment(verticalAlignment);
        //边框
        style.setBorderTop(borderStyle);
        style.setBorderBottom(borderStyle);
        style.setBorderLeft(borderStyle);
        style.setBorderRight(borderStyle);

        style.setWrapText(true);

        HSSFFont font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeightInPoints((short)fontHeight);//设置字体大小
        font.setColor(fontColor);//设置颜色
        font.setBold(bolder);

        style.setFont(font);

        return style;
    }

    /**
     * 根据Border类型获取BorderStyle对象
     * @param border
     * @return
     */
    private static BorderStyle getBorderStyle(short border) {
        if (border == 0) {
            return BorderStyle.NONE;
        }  else {
            return BorderStyle.THIN;
        }
    }

    /**
     * 根据原单元格对象获取新的HSSFCellStyle对象
     * @param srcWb
     * @param destWb
     * @param srcStyle
     * @return
     */
    private static HSSFCellStyle getStyle(HSSFWorkbook srcWb, HSSFWorkbook destWb, HSSFCellStyle srcStyle) {
        // TODO Auto-generated method stub
        HSSFCellStyle style = destWb.createCellStyle();

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //边框
        style.setBorderTop(getBorderStyle(srcStyle.getBorderTop()));
        style.setBorderBottom(getBorderStyle(srcStyle.getBorderBottom()));
        style.setBorderLeft(getBorderStyle(srcStyle.getBorderLeft()));
        style.setBorderRight(getBorderStyle(srcStyle.getBorderRight()));

        style.setWrapText(true);

        Font srcFont = srcWb.getFontAt(srcStyle.getFontIndex());

        HSSFFont font = destWb.createFont();
        font.setFontName(srcFont.getFontName());
        font.setFontHeightInPoints(srcFont.getFontHeightInPoints());//设置字体大小
        font.setColor(srcFont.getColor());//设置颜色
        font.setBold(srcFont.getBold());

        style.setFont(font);

        return style;
    }

    /**
     * 创建Excel行
     * @param wb
     * @param printRow
     * @param sheet
     * @param rowIndex
     * @return
     */
    private Row createRow(HSSFWorkbook wb, PrintRow printRow, Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int tdSize = printRow.getTds().size();

        for (int index = 0; index < tdSize; index++) {
            PrintTD printTD = printRow.getTds().get(index);
            Cell newCell = row.createCell(index);
            newCell.setCellValue(printTD.getData());
            HSSFCellStyle style = getCellStyle(wb, printTD.getStyle());
            newCell.setCellStyle(style);
        }
        return row;
    }

    /**
     * 根据CSS样式获取单元格样式对象
     * @param workbook
     * @param css
     * @return
     */
    private HSSFCellStyle getCellStyle(HSSFWorkbook workbook, String css) {
        HSSFCellStyle style = workbook.createCellStyle();

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //边框
        BorderStyle borderStyle = BorderStyle.THIN;
        if (css.indexOf("border-bottom: 1px solid") != -1) {
            style.setBorderBottom(BorderStyle.THIN);
        }
        if (css.indexOf("border-bottom: 2px solid") != -1) {
            style.setBorderBottom(BorderStyle.THICK);
        }
        if (css.indexOf("border-left: 1px solid") != -1) {
            style.setBorderLeft(BorderStyle.THIN);
        }
        if (css.indexOf("border-left: 2px solid") != -1) {
            style.setBorderLeft(BorderStyle.THICK);
        }
        if (css.indexOf("border-top: 1px solid;") != -1) {
            style.setBorderTop(BorderStyle.THIN);
        }
        if (css.indexOf("border-top: 2px solid;") != -1) {
            style.setBorderTop(BorderStyle.THICK);
        }
        if (css.indexOf("border-right: 1px solid") != -1) {
            style.setBorderRight(BorderStyle.THIN);
        }
        if (css.indexOf("border-right: 2px solid") != -1) {
            style.setBorderRight(BorderStyle.THICK);
        }

        style.setWrapText(true);

        HSSFFont font = workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)14);//设置字体大小
        font.setColor(IndexedColors.AUTOMATIC.getIndex());//设置颜色
        if (css.indexOf("font-weight: bolder") != -1) {
            font.setBold(true);
        } else {
            font.setBold(false);
        }
        style.setFont(font);

        return style;
    }
}
