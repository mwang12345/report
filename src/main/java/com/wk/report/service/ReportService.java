package com.wk.report.service;

import com.wk.report.pojo.PrintTable;
import com.wk.report.pojo.ReportIn;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public interface ReportService {

    public Map<String, List<String>> SCRIPT_MAP = new HashMap<String, List<String>>();

    public Map<String, ReportIn> REPORT_PARAM_MAP = new HashMap<String, ReportIn>();

    public Map<String, List<PrintTable>> PRINT_TABLE_MAP = new HashMap<String, List<PrintTable>>();

    public List<PrintTable> createPrintTables(HSSFWorkbook wb, ReportIn reportIn) throws Exception;
}
