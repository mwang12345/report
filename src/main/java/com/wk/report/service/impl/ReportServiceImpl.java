package com.wk.report.service.impl;

import com.wk.report.pojo.*;
import com.zdxf.report.pojo.*;
import com.wk.report.service.ReportService;
import freemarker.template.Configuration;
import freemarker.template.Template;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.StringWriter;
import java.util.*;

@Service
public class ReportServiceImpl implements ReportService {

    private static int PAGE_SIZE = 30;

    private Configuration cfg;

    private static Logger logger = LoggerFactory.getLogger(ReportService.class);

    @Value("${zdxf.report.tpl}")
    private String tplPath;

    @Value("${zdxf.report.template}")
    private String templatePath;

    @Value("${zdxf.report.genPath}")
    private String genPath;

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
    public List<String> reportPrint(ReportIn reportIn, String scriptKey) throws Exception {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(templatePath + "//" + reportIn.getFileName() + ".xls"));// 得到这个excel表格对象
        return reportPrint(wb, reportIn, scriptKey);
    }

    /**
     * 打印报表转成html内容
     * @param wb
     * @return
     */
    public List<String> reportPrint(HSSFWorkbook wb, ReportIn reportIn, String scriptKey) throws Exception{

        List<String> scripts = new ArrayList<String>();
        List<PrintTable> printTables = createPrintTables(wb, reportIn);
        PRINT_TABLE_MAP.put(scriptKey, printTables);
        for (PrintTable printTable : printTables) {
            String script = createTableHtmlScript(printTable);
            scripts.add(script);
        }
        return scripts;
    }

    /**
     * 转换Script脚本
     * @param printTable
     * @return
     * @throws Exception
     */
    private String createTableHtmlScript(PrintTable printTable) throws Exception {

        Template printTpl = cfg.getTemplate("table.tpl");
        StringWriter writer = new StringWriter();
        printTpl.process(printTable, writer);

        return writer.toString();
    }

    /**
     * 将Excel模型对象转换成表格封装对象
     * @param wb
     * @return
     */
    public List<PrintTable> createPrintTables(HSSFWorkbook wb, ReportIn reportIn) throws Exception{

        List<PrintTable> retLst = new ArrayList<PrintTable>();
        if (wb == null) {
            return retLst;
        }
        HSSFSheet sheet0 = wb.getSheetAt(0);
        Map<Integer, PrintTable> tableMap = new HashMap<>();

        int mergedNums = sheet0.getNumMergedRegions();
        List<CellRangeAddress> rangeList = new ArrayList<CellRangeAddress>();
        for(int index = 0; index < mergedNums; index++) {
            CellRangeAddress range = sheet0.getMergedRegion(index);
            rangeList.add(range);
        }

        List<PrintRow> headers = new ArrayList<>();
        Iterator<Row> rowIt = sheet0.rowIterator();
        int lastRowNum = sheet0.getLastRowNum();
        int index = 0;
        int page = 0;

        // 保存跨行跨列数据，处理相关表格边框
        List<PrintTD> rpTds = new ArrayList<>();
        boolean notFillData = true;
        int valueDataLength = 0;
        while(rowIt.hasNext()) {
            Row row = rowIt.next();
            // 从模板绘制单元格，即绘制表头信息
            PrintRow printRow = createRowFromExcel(wb, reportIn.getShowHeaderRow(), reportIn.getShowIndexs(), sheet0, row, rangeList, reportIn.getHeaderRows(), lastRowNum);
            printRow.setRowIndex(index);
            if (printRow.getTds().isEmpty()) {
                continue;
            }

            if (notFillData && isDataSample(printRow)) {

                valueDataLength = printRow.getTds().size();
                // 填充数据
                if(reportIn.getValues() == null || reportIn.getValues().length == 0) {
                    continue;
                }
                initValues(reportIn, headers.size());
                for (int rowIndex = 0; rowIndex < reportIn.getValues().length; rowIndex++) {

                    // 分页处理
                    int pageCount = page * reportIn.getPageSize();
                    if (index == pageCount) {
                        page = page + 1;
                        PrintTable table = new PrintTable();
                        table.setPage(page);
                        retLst.add(table);
                        tableMap.put(page, table);
                        index = index + headers.size() + 1;
                    } else {
                        index = index+1;
                    }

                    int curRowNum = row.getRowNum() + rowIndex;
                    PrintRow dataRow = new PrintRow();
                    dataRow.setRowIndex(index);
                    String[] cellValues = reportIn.getValues()[rowIndex].split(";");
                    List<PrintTD> cpTds = new ArrayList<>();
                    for (int index1 = 0; index1 < cellValues.length; index1++) {
                        if (!showTD(index1, reportIn.getShowIndexs())) {
                            continue;
                        }
                        PrintTD td = new PrintTD();
                        String value = cellValues[index1];
                        td.setRowIndex(index);
                        td.setColIndex(index1);
                        td.setData(value);
                        td.setStyle("border-left: 1px solid; border-top: 1px solid; text-align: center; font-size: 20px; font-family: 宋体;");

                        // Value数据合并单元格处理
                        int rpIdx = value.indexOf("rp=");
                        if (rpIdx != -1) {
                            int rowSpan = Integer.parseInt(value.substring(rpIdx+3));
                            td.setRowSpan("rowspan=" + rowSpan);
                            td.setRowSpanVal(rowSpan);
                            td.setData(value.substring(0, rpIdx - 1));
                            rpTds.add(td);

                            // 判断是否跨页，如果跨页则需要增加底部边框
                            if (td.getRowIndex() + rowSpan - 1 >= page * reportIn.getPageSize()) {
                                td.setStyle(td.getStyle() + "; border-bottom: 2px solid;");
                            }
                        }
                        int cpIdx = value.indexOf("cp=");
                        if (cpIdx != -1) {
                            int colSpan = Integer.parseInt(value.substring(cpIdx+3));
                            td.setColSpanVal(colSpan);
                            td.setColSpan("colspan=" + colSpan);
                            td.setData(value.substring(0, cpIdx - 1));
                            cpTds.add(td);
                        }
                        PrintTD colTD = handleColSpan(cpTds, td);
                        PrintTD rowTD = handleRowSpan(rpTds, td, reportIn, index1);
                        if (rowTD != null && colTD != null) {
                            dataRow.getTds().add(rowTD);
                        }
                    }
                    // 每行第一个和最后一个单元格边框加粗
                    if(!dataRow.getTds().isEmpty()) {
                        PrintTD firstTD = dataRow.getTds().get(0);
                        PrintTD lastTD = dataRow.getTds().get(dataRow.getTds().size() - 1);
                        lastTD.setStyle(lastTD.getStyle() + "; border-right: 2px solid");
                        firstTD.setStyle(firstTD.getStyle() + "; border-left: 2px solid");
                    }
                    if (curRowNum < reportIn.getHeaderRows()) {
                        headers.add(dataRow);
                    } else if (printRow.isHeader()) {
                        headers.add(dataRow);
                    } else {
                        tableMap.get(page).getRows().add(dataRow);
                    }
                }
                notFillData = false;
                break;
            } else {
                // 如果是Excel模板标题
                if(!printRow.getTds().isEmpty()) {
                    PrintTD lastTD = printRow.getTds().get(printRow.getTds().size() - 1);
                    PrintTD firstTD = printRow.getTds().get(0);
                    boolean leftBolder = true;
                    boolean rightBolder = true;
                    for (CellRangeAddress rangeAddress: rangeList) {
                        if (rangeAddress.getLastColumn() == firstTD.getColIndex() - 1) {
                           if (rangeAddress.getLastRow() >= firstTD.getRowIndex()
                                   && rangeAddress.getFirstRow() <= firstTD.getRowIndex()) {
                                leftBolder = false;
                                break;
                           }
                        }
                        if (rangeAddress.getFirstRow() == firstTD.getColIndex() + 1) {
                            if (rangeAddress.getLastRow() >= firstTD.getRowIndex()
                                    && rangeAddress.getFirstRow() <= firstTD.getRowIndex()) {
                                rightBolder = false;
                                break;
                            }
                        }
                    }
                    if (leftBolder && firstTD.getBorderLeft() != 0) {
                        firstTD.setStyle(firstTD.getStyle() + "; border-left: 2px solid");
                    }
                    if (rightBolder && lastTD.getBorderRight() != 0) {
                        lastTD.setStyle(lastTD.getStyle() + "; border-right: 2px solid");
                    }
                }
            }
            // 分页处理
            // 参数无法传Integer值，所以没有封装函数
            int pageCount = page * reportIn.getPageSize();
            if (index == pageCount) {
                page = page + 1;
                PrintTable table = new PrintTable();
                table.setPage(page);
                retLst.add(table);
                tableMap.put(page, table);
                index = index + headers.size() + 1;
            } else {
                index = index+1;
            }
            if (row.getRowNum() < reportIn.getHeaderRows()) {
                headers.add(printRow);
            } else if (printRow.isHeader()) {
                headers.add(printRow);
            } else {
                tableMap.get(page).getRows().add(printRow);
            }
        }
        for (Map.Entry<Integer, PrintTable> entry : tableMap.entrySet()) {
            PrintTable table = entry.getValue();
            table.setHeaders(headers);
        }
        // 设置总页数
        int totalPage = retLst.size();
        // 最后一条如果是个空单元格，则删除，防止末尾底线变形
        List<PrintRow> lastRows = retLst.get(retLst.size() - 1).getRows();
        PrintRow lastedRow = lastRows.get(lastRows.size() - 1);
        if (lastedRow.getTds() != null && StringUtils.isEmpty(lastedRow.getTds().get(0).getData())) {
            lastRows.remove(lastedRow);
        }

        for (PrintTable table : retLst) {
            // 每页开头加粗
            addTopBorder(table.getHeaders());

            // 每页末尾底线加粗
            table.setTotalPage(totalPage);
            table.setPageDesc("第" + table.getPage() + "页，共" + totalPage + "页");
            // 设置每页表格底部边框
            if (!table.getRows().isEmpty()) {
                PrintRow lastRow = table.getRows().get(table.getRows().size() - 1);
                for (PrintTD td : lastRow.getTds()) {
                    td.setStyle(td.getStyle() + "; border-bottom: 2px solid");
                }
            } else {
                // 如果数据没有加底部边框，则在末页的表头添加边框。
                int headerSize = table.getHeaders().size();
                for(int rowIndex = 0; rowIndex < headerSize; rowIndex++) {
                    PrintRow printRow = table.getHeaders().get(rowIndex);
                    for (PrintTD td : printRow.getTds()) {
                        if(td.getRowSpanVal() != null) {
                            if (td.getRowIndex() + td.getRowSpanVal() - 1 == headerSize) {
                                td.setStyle(td.getStyle() + "; border-bottom: 2px solid;");
                            }
                        }
                        if (td.getRowIndex() == headerSize) {
                            td.setStyle(td.getStyle() + "; border-bottom: 2px solid;");
                        }
                    }
                }
            }
        }

        insertCustomerHeader(headers, reportIn, valueDataLength);
        return retLst;
    }

    /**
     * 添加顶部边框
     * @param rows
     */
    private void addTopBorder(List<PrintRow> rows) {
        if (rows.isEmpty()) {
            return;
        }
        Integer rowIndex = -1;
        int size = rows.size();
        out: for (int index = 0; index < size; index++) {
            PrintRow row = rows.get(index);
            for (PrintTD td : row.getTds()) {
                if (td.getStyle().indexOf("border-top: 2px solid") != -1) {
                    return;
                }
                if (td.getStyle().indexOf("border-top: 1px solid") != -1) {
                    rowIndex = index;
                    break out;
                }
            }
        }
        if (rowIndex != -1) {
            PrintRow row = rows.get(rowIndex);
            for (PrintTD td : row.getTds()) {
                td.setStyle(td.getStyle() + "; border-top: 2px solid !important");
            }
        }
    }

    /**
     * 在首行标头之下，插入自定义表头
     */
    private void insertCustomerHeader(List<PrintRow> headers, ReportIn reportIn, int colSpan) {
        HeaderValue[] headerValues = reportIn.getHeaderValues();
        if (headerValues == null) {
            return;
        }
        List<PrintRow> headerValueRows = new ArrayList<>();
        int rowIndex = 0;
        for (HeaderValue headerValue : headerValues) {
            PrintRow printRow = new PrintRow();
            PrintTD labelTD = new PrintTD();
            labelTD.setData(headerValue.getLabel());
            labelTD.setStyle("font-family: 宋体; font-size: 20px;");
            printRow.getTds().add(labelTD);
            PrintTD valueTD = new PrintTD();
            valueTD.setData(headerValue.getValue());
            valueTD.setStyle("font-family: 宋体; font-size: 20px;");
            valueTD.setColSpan("colspan='" + colSpan + "'");
            printRow.getTds().add(valueTD);
            // 设置插入自定义标题左右边框
            labelTD.setStyle(labelTD.getStyle() + "; border-left: 2px solid");
            valueTD.setStyle(valueTD.getStyle() + "; border-right: 2px solid");
            // 设置自定义标题第一行头边框
            if (rowIndex == 0) {
                labelTD.setStyle(labelTD.getStyle() + "; border-top: 1px solid");
                valueTD.setStyle(valueTD.getStyle() + "; border-top: 1px solid");
            }
            headerValueRows.add(printRow);
            rowIndex++;
        }
        headers.addAll(1, headerValueRows);
    }

    /**
     * 判断是否NA单元格
     * @param printRow
     * @return
     */
    private boolean isDataSample(PrintRow printRow) {
        for (PrintTD printTD : printRow.getTds()) {
            if (printTD.getData().equals("NA")) {
                return true;
            }
        }
        return false;
    }

    /**
     * 判断是否显示单元格
     * @param index
     * @param showIndexs
     * @return
     */
    public static boolean showTD (int index, Integer[] showIndexs) {
        if (showIndexs == null) {
            return true;
        }
        for (Integer showIndex : showIndexs) {
            if (showIndex.intValue() == index) {
                return true;
            }
        }
        return false;
    }

    /**
     * 新建Table的行Row，读取Wb模板数据
     * @param wb
     * @param sheet
     * @param row
     * @param rangeList
     * @param headerRowCount
     * @param lastRowNum
     * @return
     */
    private PrintRow createRowFromExcel(HSSFWorkbook wb, Integer showHeaderRow, Integer[] showColIndexs, HSSFSheet sheet, Row row, List<CellRangeAddress> rangeList, int headerRowCount, int lastRowNum) {

        if (row == null) {
            return null;
        }
        PrintRow printRow = new PrintRow();
        Iterator<Cell> cellIt = row.cellIterator();
        int cellIndex = -1;
        while(cellIt.hasNext()) {
            cellIndex = cellIndex + 1;
            Cell cell = cellIt.next();
            int colIndex = cell.getColumnIndex();
            int rowIndex = cell.getRowIndex();

            PrintTD td = new PrintTD();
            td.setRowIndex(rowIndex);
            td.setColIndex(colIndex);
            // 合并表头单元格处理
            if (handleExcelMerge(row, printRow, td, rangeList, headerRowCount, lastRowNum)) {
                if (!showTD(cellIndex, showColIndexs)) {
                    continue;
                }
                printRow.getTds().add(td);
            }

            String value = cell.toString();
            td.setRowIndex(cell.getRowIndex());
            td.setColIndex(colIndex);
            td.setStyle(parseStyle(wb, cell));
            td.setCls("");
            // 处理表头斜线
            if(value.indexOf("\\\\") != -1) {
                StringBuilder sb = new StringBuilder();
                sb.append("<span style=\"float:left;margin-top:10px; margin-left: 5px\">" + value.split("\\\\")[0] + "</span>\n" +
                        "\t<span style=\"float:right;margin-top:-2px; margin-right: 5px\">" + value.split("\\\\")[2] + "</span>");
                td.setData(sb.toString());
                td.setCls("lineTd");
            } else {
                td.setData(cell.toString());
            }

            int width = sheet.getColumnWidth(cell.getColumnIndex());

            // 绘制标题单元格
            CellStyle cellStyle = cell.getCellStyle();

            td.setBorderLeft(cellStyle.getBorderLeft());
            td.setBorderRight(cellStyle.getBorderRight());
            td.setBorderTop(cellStyle.getBorderTop());
            td.setBorderBottom(cellStyle.getBorderBottom());

            td.setWidth(width);
            printRow.setRowWidth(printRow.getRowWidth() + width);

            td.setStyle(td.getStyle() + String.format("border-top: %spx solid; ", td.getBorderTop() + ""));
            td.setStyle(td.getStyle() + String.format("border-left: %spx solid; ", td.getBorderLeft() + ""));
        }

        // 处理宽度，等比例缩放
        for (PrintTD printTD : printRow.getTds()) {
            // 如果有合并单元格则不作处理
            if (!StringUtils.isEmpty(printTD.getColSpan())) {
                break;
            }
            double styleWidth = printTD.getWidth() / printRow.getRowWidth() * 100;
            if (styleWidth == 100) {
                break;
            }
            printTD.setStyle(printTD.getStyle() + String.format(" width:%s", styleWidth + "%;"));
        }

        if(!printRow.getTds().isEmpty()) {
            PrintTD lastTD = printRow.getTds().get(printRow.getTds().size() - 1);
            lastTD.setStyle(lastTD.getStyle() + "; border-right: 2px solid");
        }
        return printRow;
    }

    /**
     * 解析样式
     * @param wb
     * @param cell
     * @return
     */
    private String parseStyle(HSSFWorkbook wb, Cell cell) {

        if (cell == null) {
            return null;
        }
        StringBuffer style = new StringBuffer();
        style.append("font-family: 宋体;");
        CellStyle cellStyle = cell.getCellStyle();

        HSSFFont font = wb.getFontAt(cellStyle.getFontIndex());
        style.append(String.format("font-size:%s;", font.getFontHeightInPoints() + "px"));

        if (font.getBold()) {
            style.append(" font-weight:bolder;");
        }
        int alignment = cellStyle.getAlignment();
        if(alignment == 1) {
            style.append("text-align: left");
        } else if (alignment == 2) {
            style.append("text-align: center;");
        } else if (alignment == 3) {
            style.append("text-align: right;");
        }
        style.append("padding: 5px 0;");
        return style.toString();
    }

    /**
     * 处理合并单元格
     * @param row
     * @param printRow
     * @param td
     * @param rangeList
     * @param headerRowCount
     * @param lastRowNum
     * @return
     */
    private boolean handleExcelMerge(Row row, PrintRow printRow, PrintTD td, List<CellRangeAddress> rangeList,int headerRowCount, int lastRowNum) {

        for (CellRangeAddress range : rangeList) {
            int startRowIndex = range.getFirstRow();
            int startColIndex = range.getFirstColumn();
            int endRowIndex = range.getLastRow();
            int endColIndex = range.getLastColumn();
            if (startRowIndex == td.getRowIndex()
                    && startColIndex == td.getColIndex()){

                int colSpan = endColIndex - startColIndex + 1;
                td.setColSpan("colspan='" + colSpan + "'");
                td.setColSpanVal(colSpan);;

                int rowSpan = endRowIndex - startRowIndex + 1;
                td.setRowSpan("rowspan='" + rowSpan + "'");
                td.setRowSpanVal(rowSpan);

                if (colSpan > 1) {
                    int rightBorder = row.getCell(endColIndex).getCellStyle().getBorderRight();
                    if (rightBorder != 1) {
                        td.setStyle(td.getStyle() + String.format("border-right: %spx solid; ", rightBorder + ""));
                    }
                }
            }
            if (startRowIndex == td.getRowIndex()) {
                if (startColIndex < td.getColIndex()
                        && endColIndex >= td.getColIndex()){
                    return false;
                }
                if (startRowIndex != endRowIndex &&
                        endRowIndex + headerRowCount == lastRowNum) {
                    printRow.setHeader(true);
                }
            }
            if (startColIndex == endColIndex
                    && startColIndex == td.getColIndex()
                    && startRowIndex < td.getRowIndex()
                    && endRowIndex >= td.getRowIndex()) {
                return false;
            }
        }
        return true;
    }

    /**
     * 处理跨行单元格边框
     * @param rowSpans
     * @param curTD
     */
    private PrintTD handleRowSpan(List<PrintTD> rowSpans, PrintTD curTD, ReportIn reportIn, int valueIdx) {
        if (rowSpans.isEmpty()) {
            return curTD;
        }
        for (PrintTD printTD : rowSpans) {
            int rpBegin = printTD.getRowIndex();
            int rowSpan = Integer.parseInt(printTD.getRowSpan().replace("rowspan=", ""));
            int rpEnd = printTD.getRowIndex() + rowSpan;

            if (curTD.getRowIndex() > rpBegin && curTD.getRowIndex() < rpEnd) {
                // 如果在它右边相邻
                if (curTD.getColIndex() == printTD.getColIndex()) {
                    return null;
                }
                // 如果在它右边相邻
                if (curTD.getColIndex() == printTD.getColIndex() + 1) {
                    curTD.setStyle(curTD.getStyle() + " ;border-left: 1px solid !important; font-size: 20px; font-family: 宋体;");
                }
            }
        }
        return curTD;
    }

    /**
     * 处理跨列单元格边框
     * @param colSpans
     * @param curTD
     */
    private PrintTD handleColSpan(List<PrintTD> colSpans, PrintTD curTD) {
        if (colSpans.isEmpty()) {
            return curTD;
        }
        for (PrintTD printTD : colSpans) {
            int cpBegin = printTD.getColIndex();
            int colSpan = Integer.parseInt(printTD.getColSpan().replace("colspan=", ""));
            int cpEnd = printTD.getColIndex() + colSpan;
            if (curTD.getColIndex() > cpBegin && curTD.getColIndex() < cpEnd) {
                return null;
            }
        }
        return curTD;
    }

    /**
     * 初始化数据，主要处理跨页合并需求
     * @param reportIn
     * @param headerSize
     * @return
     */
    private String[] initValues(ReportIn reportIn, int headerSize) {
        String[] values = reportIn.getValues();
        String[] ret = new String[reportIn.getValues().length];
        if (values == null || values.length == 0) {
            return values;
        }
        int curPage = 0;
        int curRowIdx = 1;
        for (int rowIdx = 0; rowIdx < values.length; rowIdx++) {
            if (curRowIdx % reportIn.getPageSize() == 1) {
                curPage = curPage + 1;
                curRowIdx = curRowIdx + headerSize;
            }

            String value = values[rowIdx];
            String[] cells = value.split(";");
            StringJoiner sb = new StringJoiner(";");

            String cell0 = cells[0];
            int rpIdx = cell0.indexOf("rp=");
            if (rpIdx != -1) {
                int rowSpan = Integer.parseInt(cell0.substring(rpIdx + 3));
                String cellData = cell0.substring(0, rpIdx - 1);
                int page1 = curRowIdx / reportIn.getPageSize();
                if (curRowIdx % reportIn.getPageSize() != 0) {
                    page1 = page1 + 1;
                }
                int page2 = (curRowIdx + rowSpan - 1) / reportIn.getPageSize();
                if ((curRowIdx + rowSpan - 1) % reportIn.getPageSize() != 0) {
                    page2 = page2 + 1;
                }
                // 如果跨页
                if (page1 != page2) {
                    int rowSpan1 = curPage * reportIn.getPageSize() - curRowIdx + 1;
                    int rowSpan2 = rowSpan - rowSpan1;
                    int rowSpanEnd = rowIdx + rowSpan1;
                    System.out.println(rowIdx);
                    values[rowIdx] = changeRowSpan(values[rowIdx], cellData, 0, rowSpan1);
                    if (rowSpanEnd < values.length) {
                        values[rowSpanEnd] = changeRowSpan(values[rowSpanEnd], cellData, 0, rowSpan2);
                    }
                }
            }
            curRowIdx = curRowIdx + 1;
        }
        return values;
    }

    private String changeRowSpan(String value, String cellData, int colIdx, int newRowSpan) {
        String[] values = value.split(";");
        StringJoiner sb = new StringJoiner(";");
        for (int index = 0; index < values.length; index++) {
            if (index == colIdx) {
                sb.add(cellData + "/rp=" + newRowSpan);
            } else {
                sb.add(values[index]);
            }
        }
        return sb.toString();
    }
}
