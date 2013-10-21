package org.bzewdu.tools.chart.soffice;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Config {
    private static String configFilePath = "config" + File.separator + "config.xml";
    private static Config ourInstance = null;//new Config();
    public static boolean debug = false;
    private Document doc = null;

    public List<sheetConfig> getSheetConfigs() {
        return sheetConfigs;
    }

    public void setSheetConfigs(List<sheetConfig> sheetConfigs) {
        this.sheetConfigs = sheetConfigs;
    }

    private List<sheetConfig> sheetConfigs = new ArrayList<sheetConfig>();

    public static Config getInstance() {
        if (ourInstance == null) {
            ourInstance = new Config();
            ourInstance.init();
        }
        return ourInstance;
    }

    private Config() {
        SAXBuilder builder = new SAXBuilder();
        try {
            doc = builder.build(configFilePath);
        } catch (JDOMException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void init() {
        Element root = doc.getRootElement();
        List<Element> sheetConfigElementList = root.getChildren();
        String name = null;
        sheetConfig sc = null;

        for (Element sheetConfigElement : sheetConfigElementList) {
            sc = new sheetConfig();
            name = sheetConfigElement.getName();
            if (name.equalsIgnoreCase("sheet")) {
                List<Element> sheetConfigItemsList = sheetConfigElement.getChildren();
                for (Element e : sheetConfigItemsList) {
                    name = e.getName();
                    if (name.equalsIgnoreCase("builds")) {
                        List<Element> buildList = e.getChildren();
                        for (Element buildElement : buildList)
                            sc.getBuildMap().put(buildElement.getAttributeValue("name"), buildElement.getAttributeValue("path"));
                    } else if (name.equalsIgnoreCase("benchmarks"))
                        sc.setBenchList(e.getAttributeValue("name").split(","));
                        /*else if (name.equalsIgnoreCase("machine"))
                        sc.setPlatform(e.getAttributeValue("os") + '.' + e.getAttributeValue("name"));
                    else if (name.equalsIgnoreCase("baseUrls"))
                        sc.setBaseUrlStr(e.getText().split(","));*/
                    else if (name.equalsIgnoreCase("document")) {
                        sc.setChartTitle(e.getChild("chart").getAttributeValue("title"));
                        sc.setChartType(e.getChild("chart").getAttributeValue("type"));
                        sc.setDocumentTitle(e.getChild("title").getText());
                    }
                    //else if (name.equalsIgnoreCase("runTypes"))
                    //    sc.setRunTypes(e.getText().split(","));
                    else
                        throw new NullPointerException(name);
                }
                sheetConfigs.add(sc);
            }
        }
    }
}
