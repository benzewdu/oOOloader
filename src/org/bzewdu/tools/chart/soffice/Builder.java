package org.bzewdu.tools.chart.soffice;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.XCell;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XNumberFormatsSupplier;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.*;

public class Builder {
    private static final int ROW_OFFSET = 2;
    private int nNumberFormat;

    public Builder(XSpreadsheetDocument _theSpreadsheetDocument) {
        XNumberFormatsSupplier xNumberFormatsSupplier = (XNumberFormatsSupplier) UnoRuntime.queryInterface(XNumberFormatsSupplier.class, _theSpreadsheetDocument);

        Locale aLocale = new Locale();
        com.sun.star.util.XNumberFormats xNumberFormats = xNumberFormatsSupplier.getNumberFormats();
        com.sun.star.util.XNumberFormatTypes xNumberFormatTypes =
                (com.sun.star.util.XNumberFormatTypes) UnoRuntime.queryInterface(com.sun.star.util.XNumberFormatTypes.class, xNumberFormats);
        nNumberFormat = xNumberFormatTypes.getStandardFormat(com.sun.star.util.NumberFormat.PERCENT, aLocale);
    }

    private void populateSheet(sheetConfig sc, XSpreadsheet xSheet) {

        List<String> lb = Arrays.asList(sc.getBenchList());
        List<Object> lm = Arrays.asList(sc.getBuildMap().keySet().toArray());

        populateHeaders(sc, lb, xSheet, lm);

        sc.setCellRange(((char) (ROW_OFFSET + lb.size() + 65) + "" + 1) + ":" + ((char) ((ROW_OFFSET + (2 * lb.size())) + 65) + "" + (lm.size() + 1)));

        formatCells(xSheet, ROW_OFFSET + lb.size() + 1, 1, ROW_OFFSET + (2 * lb.size()), (lm.size() + 1));

        populateCells(sc, lb, lm, xSheet);
    }

    private void populateCells(sheetConfig sc,
                               List<String> lb,
                               List<Object> lm,
                               XSpreadsheet xSheet) {

        List<BigDecimal> scores = null;
        for (String benchmark : sc.getBenchList()) {
            for (String buildName : sc.getBuildMap().keySet()) {
                scores = getResultData(sc, buildName, benchmark);
                if (scores != null) {

                    insertIntoCell(
                            lb.indexOf(benchmark) + 1,
                            lm.indexOf(buildName) + 1,
                            "" + Helper.bd_getMean(scores, false),
                            xSheet,
                            "V");

                    String baseline = (char) ((lb.indexOf(benchmark)) + 1 + 65) + "" + 2;
                    String specimen = (char) ((lb.indexOf(benchmark)) + 1 + 65) + "" + (lm.indexOf(buildName) + 2);

                    insertIntoCell(
                            lb.indexOf(benchmark) + 1 + ROW_OFFSET + lb.size(),
                            lm.indexOf(buildName) + 1,

                            "=1-((" + baseline + "-" + specimen + ")/" + baseline + ")",
                            xSheet,
                            "");
                }
            }
        }
    }

    private List<BigDecimal> getResultData(sheetConfig sc, String buildName, String benchmark) {
        List<BigDecimal> scores = new ArrayList<BigDecimal>();
        String buildScorePath = sc.getBuildMap().get(buildName);

        BigDecimal score = getScoreFromResultFile(buildScorePath + File.separator + "results." + benchmark + File.separator + "results." + benchmark);
        if (score != null)
            scores.add(score);
        return scores;
    }

    private static synchronized BigDecimal getScoreFromResultFile(String resultFileName) {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(resultFileName);
            Properties scoreProps = new Properties();
            scoreProps.load(fis);
            return new BigDecimal(scoreProps.getProperty("score"));
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null)
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
        }
        return null;
    }

    private void populateHeaders(sheetConfig sc, List<String> lb, XSpreadsheet xSheet, List<Object> lm) {
        for (String benchmark : sc.getBenchList()) {
            insertIntoCell(lb.indexOf(benchmark) + 1, 0, benchmark, xSheet, "");
            insertIntoCell(lb.indexOf(benchmark) + 1 + ROW_OFFSET + lb.size(), 0, benchmark, xSheet, "");
        }
        for (String buildName : sc.getBuildMap().keySet()) {
            insertIntoCell(0, lm.indexOf(buildName) + 1, buildName, xSheet, "");
            insertIntoCell(ROW_OFFSET + lb.size(), lm.indexOf(buildName) + 1, buildName, xSheet, "");
        }
    }

    public void renderSheets(XSpreadsheets xSheets) {
        final Chart chart = new Chart();
        HashMap<sheetConfig, XSpreadsheet> xSheetsMap = new HashMap<sheetConfig, XSpreadsheet>();

        for (final sheetConfig sc : Config.getInstance().getSheetConfigs()) {
            xSheets.insertNewByName(sc.getDocumentTitle(), (short) 1);
            XIndexAccess oIndexSheets = (XIndexAccess) UnoRuntime.queryInterface(XIndexAccess.class, xSheets);
            try {
                xSheetsMap.put(sc, (XSpreadsheet) UnoRuntime.queryInterface(XSpreadsheet.class, oIndexSheets.getByIndex(1)));
            } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
                e.printStackTrace();
            } catch (WrappedTargetException e) {
                e.printStackTrace();
            }
        }
        List<Thread> threadList = new ArrayList<Thread>();
        Thread thread = null;
        for (final sheetConfig sc : xSheetsMap.keySet()) {
            final XSpreadsheet xSheet = xSheetsMap.get(sc);
            threadList.add(thread = new Thread() {
                public void run() {
                    populateSheet(sc, xSheet);
                    try {
                        chart.draw(xSheet, sc);
                    } catch (NoSuchElementException e) {
                        e.printStackTrace();
                    } catch (WrappedTargetException e) {
                        e.printStackTrace();
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    } catch (UnknownPropertyException e) {
                        e.printStackTrace();
                    } catch (PropertyVetoException e) {
                        e.printStackTrace();
                    } catch (com.sun.star.lang.IllegalArgumentException e) {
                        e.printStackTrace();
                    }
                }
            });
            thread.start();
        }

        try {
            for (Thread t : threadList) t.join();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

    }

    private void insertIntoCell(int CellX, int CellY, String theValue,
                                XSpreadsheet TT1, String flag) {
        XCell xCell = null;

        try {
            xCell = TT1.getCellByPosition(CellX, CellY);
        } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
            if (Config.debug) if (Config.debug) System.err.println("Could not get Cell");
            ex.printStackTrace(System.err);
        }

        if (flag.equals("V")) {
            assert xCell != null;
            xCell.setValue(new Float(theValue));
        } else {
            assert xCell != null;
            xCell.setFormula(theValue);
        }

    }

    private void formatCells(XSpreadsheet xSheet, int i, int j, int k, int l) {

        com.sun.star.table.XCellRange xCellRange = null;
        try {
            xCellRange = xSheet.getCellRangeByPosition(i, j, k, l);
        } catch (com.sun.star.lang.IndexOutOfBoundsException e) {
            e.printStackTrace();
        }

        com.sun.star.beans.XPropertySet xCellProp = (com.sun.star.beans.XPropertySet) UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, xCellRange);

        try {
            xCellProp.setPropertyValue("NumberFormat", nNumberFormat);
        } catch (UnknownPropertyException e) {
            e.printStackTrace();
        } catch (PropertyVetoException e) {
            e.printStackTrace();
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        } catch (WrappedTargetException e) {
            e.printStackTrace();
        }

    }
}
