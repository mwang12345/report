package com.wk.report.service;

import com.wk.report.pojo.ReportIn;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 导出报表
 */
public interface ExportService {

    public HSSFWorkbook export(ReportIn reportIn, String fileName) throws Exception;

    public HSSFWorkbook exportWithoutTemplate(ReportIn reportIn) throws Exception;

}
