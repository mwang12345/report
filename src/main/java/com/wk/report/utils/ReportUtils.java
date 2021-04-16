package com.wk.report.utils;

import com.wk.report.pojo.ReportIn;
import org.springframework.beans.BeanUtils;

public class ReportUtils {

    public static ReportIn clone(ReportIn reportIn) {
        ReportIn ret = new ReportIn();
        BeanUtils.copyProperties(reportIn, ret);
        ret.values = new String[reportIn.getValues().length];
        for (int index = 0; index < reportIn.values.length; index++) {
            ret.values[index] = reportIn.values[index];
        }
        return ret;
    }
}
