package org.bzewdu.tools.chart.soffice;


import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;


public class Main {

    private static XSpreadsheetDocument theSpreadsheetDocument = null;
    private static XComponent _xComponent;

    public static void main(String args[]) {
        XComponentContext xContext = null;

        try {
            xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
            if (Config.debug) System.out.println("Connected to a running office ...");
        } catch (Exception e) {
            e.printStackTrace(System.err);
            System.exit(1);
        }

        if (Config.debug) System.out.println("Opening an empty Calc document");
        openCalc(xContext);

        if (Config.debug) System.out.println("Getting spreadsheet");
        XSpreadsheets xSheets = getTheSpreadsheetDocument().getSheets();

        if (Config.debug) System.out.println("Rendering spreadsheet");
        Builder builder = new Builder(getTheSpreadsheetDocument());
        builder.renderSheets(xSheets);

        saveAndCloseDocComponent(get_xComponent(), args[0]);

        try {
            Thread.sleep(2000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        if (Config.debug) System.out.println("All Done");
        System.exit(0);
    }


    private static void openCalc(XComponentContext xContext) {
        //define variables
        XMultiComponentFactory xMCF = null;
        XComponentLoader xCLoader;
        XSpreadsheetDocument xSpreadSheetDoc = null;
        XComponent xComp = null;

        try {
            // get the service manager rom the office
            xMCF = xContext.getServiceManager();

            // create a new instance of the the desktop
            Object oDesktop = xMCF.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", xContext);

            // query the desktop object for the XComponentLoader
            xCLoader = (XComponentLoader) UnoRuntime.queryInterface(
                    XComponentLoader.class, oDesktop);

            PropertyValue[] szEmptyArgs = new PropertyValue[0];
            String strDoc = "private:factory/scalc";

            set_xComponent(xComp = xCLoader.loadComponentFromURL(strDoc, "_blank", 0, szEmptyArgs));
            setTheSpreadsheetDocument((XSpreadsheetDocument) UnoRuntime.queryInterface(
                    XSpreadsheetDocument.class, xComp));

        } catch (Exception e) {
            if (Config.debug) System.err.println(" Exception " + e);
            e.printStackTrace(System.err);
        }

    }


    private static XSpreadsheetDocument getTheSpreadsheetDocument() {
        return theSpreadsheetDocument;
    }

    private static XComponent get_xComponent() {
        return _xComponent;
    }

    private static void setTheSpreadsheetDocument(XSpreadsheetDocument __theSpreadsheetDocument) {
        theSpreadsheetDocument = __theSpreadsheetDocument;
    }

    private static void set_xComponent(XComponent __xComponent) {
        _xComponent = __xComponent;
    }


    private static void saveAndCloseDocComponent(XComponent xDoc, String storeUrl) {
        XStorable xStorable = (XStorable) UnoRuntime.queryInterface(XStorable.class, xDoc);

        PropertyValue[] storeProps = new PropertyValue[1];
        storeProps[0] = new PropertyValue();
        storeProps[0].Name = "FilterName";
        storeProps[0].Value = "StarCalc 5.0";

        try {
            xStorable.storeAsURL(storeUrl, storeProps);
        } catch (IOException e) {
            e.printStackTrace();
        }

        XCloseable xcloseable = (XCloseable) UnoRuntime.queryInterface(XCloseable.class, xStorable);

        // Closing the converted document
        if (xcloseable != null)
            try {
                xcloseable.close(false);
            } catch (CloseVetoException e) {
                e.printStackTrace();  //To change body of catch statement use File | Settings | File Templates.
            }
        else {
            // If Xcloseable is not supported (older versions,
            // use dispose() for closing the document
            XComponent xComponent = (XComponent) UnoRuntime.queryInterface(
                    XComponent.class, xStorable);
            xComponent.dispose();
        }
    }
}
