package com.wk.report.pojo;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class PrintTable {

    private String style;

    private List<PrintRow> titles = new ArrayList<PrintRow>();

    private List<PrintRow> headers = new ArrayList<PrintRow>();

    private List<PrintRow> rows = new ArrayList<PrintRow>();

    private int page = 1;

    private int totalPage;

    private String pageDesc;

}
