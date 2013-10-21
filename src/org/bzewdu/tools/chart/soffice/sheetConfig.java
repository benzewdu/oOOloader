package org.bzewdu.tools.chart.soffice;

import java.util.Map;
import java.util.TreeMap;


class sheetConfig {
    private String documentTitle;
    private String cellRange = null;


    public String getCellRange() {
        return cellRange;
    }

    public void setCellRange(String _cellRange) {
        cellRange = _cellRange;
    }

    public String[] getrunType() {
        return runType;
    }

    public void setRunTypes(String[] runType) {
        this.runType = runType;
    }

    private String[] runType;

    public Map<String, String> getBuildMap() {
        return buildMap;
    }

    public void setBuildMap(Map<String, String> buildMap) {
        this.buildMap = buildMap;
    }

    public String[] getBenchList() {
        return benchList;
    }

    public void setBenchList(String[] benchList) {
        this.benchList = benchList;
    }

    public String getPlatform() {
        return platform;
    }

    public void setPlatform(String platform) {
        this.platform = platform;
    }


    public String getChartTitle() {
        return chartTitle;
    }

    public void setChartTitle(String chartTitle) {
        this.chartTitle = chartTitle;
    }

    public String getChartType() {
        return chartType;
    }

    public void setChartType(String chartType) {
        this.chartType = chartType;
    }

    public String[] getBaseUrlStr() {
        return baseUrlStr;
    }

    public void setBaseUrlStr(String[] baseUrlStr) {
        this.baseUrlStr = baseUrlStr;
    }

    private String[] baseUrlStr = null;
    private Map<String, String> buildMap = new TreeMap<String, String>();
    private String[] benchList;
    private String platform = null;
    private String chartTitle = null;
    private String chartType = null;


    public void setDocumentTitle(String text) {
        documentTitle = text;
    }

    public String getDocumentTitle() {
        return documentTitle;
    }
}