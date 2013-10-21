package org.bzewdu.tools.chart.soffice;

import java.math.BigDecimal;
import java.util.List;

public class Helper {

    public static double bd_getMean(List<BigDecimal> data, boolean round) {
        if (data == null || data.size() <= 0)
            return (double) 0;
        if (!round)
            return bd_sum(data) / (double) data.size();
        else
            return (new BigDecimal(bd_sum(data) / (double) data.size())).setScale(3, 1).doubleValue();
    }

    private static double bd_sum(List data) {
        //double sum = 0;
        BigDecimal _sum = new BigDecimal(0.0D);
        for (int i = 0; i < data.size(); i++)
            if (data.get(i) != null)
                _sum = _sum.add((BigDecimal) data.get(i));

        return _sum.doubleValue();
    }
}
