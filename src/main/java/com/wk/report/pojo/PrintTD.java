package com.wk.report.pojo;

import lombok.Data;

@Data
public class PrintTD {

    private String style;

    private String data;

    private int rowIndex;

    private int colIndex;

    private double width;

    private double height;

    private Integer colSpanVal;

    private Integer rowSpanVal = 0;

    private String colSpan = "";

    private String rowSpan = "";

    private String cls = "";

    private short borderLeft = 0;

    private short borderTop = 0;

    private short borderRight = 0;

    private short borderBottom = 0;

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        PrintTD printTD = (PrintTD) o;

        if (this.rowIndex == ((PrintTD) o).rowIndex
                && this.colIndex == ((PrintTD) o).colIndex) {
            return true;
        }
        return false;
    }

}
