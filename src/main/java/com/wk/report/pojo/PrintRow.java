package com.wk.report.pojo;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class PrintRow {

    private String style;

    private List<PrintTD> tds = new ArrayList<PrintTD>();

    private int rowIndex;

    private double rowHeight;

    private boolean isHeader;

    private boolean title;

    private int index;

    private long headerTimeStamp;

    private double rowWidth;

}
