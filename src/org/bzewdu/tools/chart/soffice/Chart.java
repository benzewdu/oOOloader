package org.bzewdu.tools.chart.soffice;

import com.sun.star.awt.Rectangle;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.chart.ChartDataRowSource;
import com.sun.star.chart.XChartDocument;
import com.sun.star.chart.XDiagram;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.document.XEmbeddedObjectSupplier;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.*;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XInterface;

class Chart {
    public void draw(XSpreadsheet xSheet, sheetConfig sc)
            throws
            NoSuchElementException,
            WrappedTargetException,
            InterruptedException,
            UnknownPropertyException,
            PropertyVetoException,
            com.sun.star.lang.IllegalArgumentException {
        Rectangle oRect = new Rectangle();
        oRect.X = 500;
        oRect.Y = 3000;
        oRect.Width = 25000;
        oRect.Height = 11000;

        XCellRange oRange = (XCellRange) UnoRuntime.queryInterface(
                XCellRange.class, xSheet);
        XCellRange myRange = oRange.getCellRangeByName(sc.getCellRange()/*"A1:C7"*/);
        XCellRangeAddressable oRangeAddr = (XCellRangeAddressable)
                UnoRuntime.queryInterface(XCellRangeAddressable.class, myRange);
        CellRangeAddress myAddr = oRangeAddr.getRangeAddress();

        CellRangeAddress[] oAddr = new CellRangeAddress[1];
        oAddr[0] = myAddr;
        XTableChartsSupplier oSupp = (XTableChartsSupplier) UnoRuntime.queryInterface(
                XTableChartsSupplier.class, xSheet);

        XTableChart oChart = null;

        if (Config.debug) System.out.println("Insert Chart");

        XTableCharts oCharts = oSupp.getCharts();
        oCharts.addNewByName(sc.getChartTitle(), oRect, oAddr, true, true);

        // get the diagramm and Change some of the properties

        oChart = (XTableChart) (UnoRuntime.queryInterface(
                XTableChart.class, ((XNameAccess) UnoRuntime.queryInterface(
                XNameAccess.class, oCharts)).getByName(sc.getChartTitle())));
        XEmbeddedObjectSupplier oEOS = (XEmbeddedObjectSupplier)
                UnoRuntime.queryInterface(XEmbeddedObjectSupplier.class, oChart);
        XInterface oInt = oEOS.getEmbeddedObject();
        XChartDocument xChart = (XChartDocument) UnoRuntime.queryInterface(
                XChartDocument.class, oInt);

        XDiagram oDiag = (XDiagram) xChart.getDiagram();
        XPropertySet oCPS = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, oDiag);


        oCPS.setPropertyValue("DataRowSource", ChartDataRowSource.ROWS);

        Thread.sleep(200);
        XPropertySet oTPS = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, xChart.getTitle());
        oTPS.setPropertyValue("String", sc.getChartTitle());

        if (Config.debug) System.err.println("-------------------------------------");

    }
}
