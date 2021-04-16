package com.wk.report.controller;

import com.wk.report.pojo.ReportIn;
import com.wk.report.service.ExportService;
import com.wk.report.service.impl.ReportServiceImpl;
import com.wk.report.utils.FileUtils;
import com.wk.report.utils.ReportUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

@Controller
@RequestMapping(value="/report")
@Slf4j
public class PrintController {

    @Autowired
    private ReportServiceImpl reportService;

    @Autowired
    private ExportService exportService;

    @Value("${zdxf.report.genPath}")
    private String genPath;

    @RequestMapping(value = "/create",method = {RequestMethod.POST})
    @ResponseBody
    public String create(Model model, @RequestBody ReportIn reportIn) throws Exception{
        if (reportIn.getPageSize() == null){
            reportIn.setPageSize(1000);
        }
        if (reportIn.getHeaderRows() == null) {
            reportIn.setHeaderRows(0);
        }
        String scriptKey = new Date().getTime() + "";
        reportService.REPORT_PARAM_MAP.put(scriptKey, ReportUtils.clone(reportIn));
        List<String> scripts = reportService.reportPrint(reportIn, scriptKey);
        reportService.SCRIPT_MAP.put(scriptKey, scripts);
        return scriptKey;
    }

    @RequestMapping(value = "/draw/{scriptKey}",method = {RequestMethod.GET})
    public String draw(Model model, @PathVariable(name = "scriptKey", required = true) String scriptKey, @RequestParam(value = "pageNo", required = false) Integer pageNo) {
        if (pageNo == null) {
            pageNo = 0;
        }
        List<String> scripts = reportService.SCRIPT_MAP.get(scriptKey);
        if (scripts != null && pageNo > scripts.size() -1) {
            pageNo = scripts.size() - 1;
        }
        String scriptValue = reportService.SCRIPT_MAP.get(scriptKey).get(pageNo);
        model.addAttribute("table", scriptValue);
        model.addAttribute("scriptKey", scriptKey);
        model.addAttribute("pageNo", pageNo);
        model.addAttribute("fileName", "");
        return "myReport";
    }

    @RequestMapping(value = "/draw/export/download/{scriptKey}", method = {RequestMethod.GET})
    public void exportDownload(HttpServletResponse response, @PathVariable(name = "scriptKey", required = true) String scriptKey) throws Exception {
        ReportIn reportIn = reportService.REPORT_PARAM_MAP.get(scriptKey);
        String exportName = reportIn.getExportName();
        if (StringUtils.isEmpty(exportName)) {
            exportName = reportIn.getFileName();
        }
        String fileName = exportName + "_" + scriptKey + ".xls";
        FileInputStream in = new FileInputStream(genPath + "/" + fileName);
        FileUtils.downloadFile(response, in, fileName);
    }

    /**
     * Excel报表导出
     * @param model
     * @param request
     * @param reportIn
     * @return
     * @throws Exception
     */
    @RequestMapping(value = "/draw/exportExcel",method = {RequestMethod.POST})
    @ResponseBody
    public String exportExcel(Model model, HttpServletRequest request,
                              @RequestBody ReportIn reportIn) throws Exception{
        FileOutputStream fos = null;
        try{
            String exportName = reportIn.getFileName();
            if (StringUtils.isEmpty(exportName)) {
                exportName = reportIn.getFileName();
            }
            String fileName = exportName + "_" + new Date().getTime() + ".xls";
            HSSFWorkbook wbook = exportService.export(reportIn, fileName);  //测试下载用，所以这里excel内容直接写死

            File file = new File(genPath + "//" + fileName);
            fos = new FileOutputStream(file);
            wbook.write(fos);// 写文件
            return file.getAbsolutePath();
        } catch (IOException e1) {
            throw e1;
        } finally{
            try {
                if(null != fos){
                    fos.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    @RequestMapping(value = "/draw/exportWithoutTemplate",method = {RequestMethod.POST})
    @ResponseBody
    public String exportWithoutTemplate(Model model, HttpServletRequest request, HttpServletResponse response,
                         @RequestBody ReportIn reportIn) throws Exception{
        FileOutputStream fos = null;
        try{
            String scriptKey = new Date().getTime() + "";
            String fileName = reportIn.getFileName() + "_" + scriptKey + ".xls";
            HSSFWorkbook wbook = exportService.exportWithoutTemplate(reportIn);  //测试下载用，所以这里excel内容直接写死

            File file = new File(genPath + "//" + fileName);
            fos = new FileOutputStream(file);
            wbook.write(fos);// 写文件
            return fileName;
        } catch (IOException e1) {
            log.error(e1.getMessage());
            return null;
        } finally{
            try {
                if(null != fos){
                    fos.close();
                }
            } catch (IOException e) {
                log.error(e.getMessage());
            }
        }
    }

    @RequestMapping(value = "/draw/export/{scriptKey}",method = {RequestMethod.GET})
    @ResponseBody
    public String export(Model model, HttpServletRequest request, HttpServletResponse response,
                       @PathVariable(name = "scriptKey", required = true) String scriptKey) throws Exception{
        String scriptValue = reportService.SCRIPT_MAP.get(scriptKey).get(0);
        FileOutputStream fos = null;
        try{
            ReportIn reportIn = reportService.REPORT_PARAM_MAP.get(scriptKey);
            String exportName = reportIn.getExportName();
            if (StringUtils.isEmpty(exportName)) {
                exportName = reportIn.getFileName();
            }
            String fileName = exportName + "_" + scriptKey + ".xls";
            HSSFWorkbook wbook = exportService.export(reportIn, fileName);  //测试下载用，所以这里excel内容直接写死

            File file = new File(genPath + "//" + fileName);
            fos = new FileOutputStream(file);
            wbook.write(fos);// 写文件
            return fileName;
        } catch (IOException e1) {
            log.error(e1.getMessage());
            return null;
        } finally{
            try {
                if(null != fos){
                    fos.close();
                }
            } catch (IOException e) {
                log.error(e.getMessage());
            }
        }
    }

    @RequestMapping(value = "/draw/print/{scriptKey}",method = {RequestMethod.GET})
    @ResponseBody
    public List<String> print(@PathVariable(name = "scriptKey", required = true) String scriptKey) {
        return reportService.SCRIPT_MAP.get(scriptKey);
    }

}
