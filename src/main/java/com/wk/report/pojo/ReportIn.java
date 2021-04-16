package com.wk.report.pojo;

import lombok.Data;

@Data
public class ReportIn {

    public String fileName;

    public Integer headerRows = 1;

    public String[] values;

    public Integer[] showIndexs;

    public HeaderValue[] headerValues;

    private Integer showHeaderRow;

    public Integer pageSize;

    private String tplPath;

    private String templatePath;

    private String exportName;

    private String genPath;
}
