/*
 * 文件名：ConverWordByMsOffice.java
 * 版权：Copyright by www.chinauip.com
 * 描述：
 * 修改人：Administrator
 * 修改时间：2017年10月12日
 * 跟踪单号：
 * 修改单号：
 * 修改内容：
 */

package com.test.docTodpf;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * 
 * 使用 jacob调用ms office应用 word 转 pdf
 * 〈功能详细描述〉
 * @author suwy
 * @version 2017年10月12日
 * @see ConverWordByMsOffice
 * @since
 */
public class ConverWordByMsOffice
{

    private ActiveXComponent oleComponent = null;
    private Dispatch activeDoc = null;
    private final static String APP_ID = "Word.Application";

    // Constants that map onto Word's WdSaveOptions enumeration and that
    // may be passed to the close(int) method
    public static final int DO_NOT_SAVE_CHANGES = 0;
    public static final int PROMPT_TO_SAVE_CHANGES = -2;
    public static final int SAVE_CHANGES = -1;

    // These constant values determine whether or not tha application
    // instance will be displyed on the users screen or not.
    public static final boolean VISIBLE = true;
    public static final boolean HIDDEN = false;

    /**
     * Create a new instance of the JacobWordSearch class using the following
     * parameters.
     *
     * @param visibility A primitive boolean whose value will determine whether
     *                   or not the Word application will be visible to the user. Pass true
     *                   to display Word, false otherwise.
     */
    public ConverWordByMsOffice(boolean visibility) {
        System.out.println("启动----");
        this.oleComponent = new ActiveXComponent(ConverWordByMsOffice.APP_ID);
        this.oleComponent.setProperty("Visible", new Variant(visibility));
    }

    /**
     * Open ana existing Word document.
     *
     * @param docName An instance of the String class that encapsulates the
     *                path to and name of a valid Word file. Note that there are a few
     *                limitations applying to the format of this String; it must specify
     *                the absolute path to the file and it must not use the single forward
     *                slash to specify the path separator.
     */
    public void openDoc(String docName) {
        System.out.println("打开--"+docName);
        Dispatch disp = null;
        Variant var = null;
        // First get a Dispatch object referencing the Documents collection - for
        // collections, think of ArrayLists of objects.
        var = Dispatch.get(this.oleComponent, "Documents");
        disp = var.getDispatch();
        // Now call the Open method on the Documents collection Dispatch object
        // to both open the file and add it to the collection. It would be possible
        // to open a series of files and access each from the Documents collection
        // but for this example, it is simpler to store a reference to the
        // active document in a private instance variable.
        var = Dispatch.call(disp, "Open", docName);
        this.activeDoc = var.getDispatch();
        System.out.println("打开结束--"+docName);
    }

    /**
     * There is more than one way to convert the document into PDF format, you
     * can either explicitly use a FileConvertor object or call the
     * ExportAsFixedFormat method on the active document. This method opts for
     * the latter and calls the ExportAsFixedFormat method passing the name
     * of the file along with the integer value of 17. This value maps onto one
     * of Word's constants called wdExportFormatPDF and causes the application
     * to convert the file into PDF format. If you wanted to do so, for testing
     * purposes, you could add another value to the args array, a Boolean value
     * of true. This would open the newly converted document automatically.
     *
     * @param filename
     */
    public void publishAsPDF(String filename) {
        // The code to expoort as a PDF is 17
        //Object args = new Object{filename, new Integer(17), new Boolean(true)};
        Object[] args =  {
            filename, new Integer(17)
        } ;
        System.out.println("转换--"+filename);
        Dispatch.call(this.activeDoc, "SaveAs", args);
    }

    /**
     * Called to close the active document. Note that this method simply
     * calls the overloaded closeDoc(int) method passing the value 0 which
     * instructs Word to close the document and discard any changes that may
     * have been made since the document was opened or edited.
     */
    public void closeDoc() {
        this.closeDoc(ConverWordByMsOffice.DO_NOT_SAVE_CHANGES);
    }

    /**
     * Called to close the active document. It is possible with this overloaded
     * version of the close() method to specify what should happen if the user
     * has made changes to the document that have not been saved. There are three
     * possible value defined by the following manifest constants;
     * DO_NOT_SAVE_CHANGES - Close the document and discard any changes
     * the user may have made.
     * PROMPT_TO_SAVE_CHANGES - Display a prompt to the user asking them
     * how to proceed.
     * SAVE_CHANGES - Save the changes the user has made to the document.
     *
     * @param saveOption A primitive integer whose value indicates how the close
     *                   operation should proceed if the user has made changes to the active
     *                   document. Note that no checks are made on the value passed to
     *                   this argument.
     */
    public void closeDoc(int saveOption) {
        Object[] args = {new Integer(saveOption)};
        Dispatch.call(this.activeDoc, "Close", args);
    }

    /**
     * Called once processing has completed in order to close down the instance
     * of Word.
     */
    public void quit() {
        Dispatch.call(this.oleComponent, "Quit");
    }
    
    
    public static void main(String[] args)
    {
        ConverWordByMsOffice obj =  new ConverWordByMsOffice(false);
        obj.openDoc("E:/市局电子化/准予设立登记通知书确认版.docx");
        obj.publishAsPDF("F:/text.pdf");
        obj.closeDoc();
        obj.quit();
    }
    

}
