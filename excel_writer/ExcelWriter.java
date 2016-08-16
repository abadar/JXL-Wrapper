package excel_writer;

import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

import jxl.*;
import jxl.write.*;
import jxl.write.Number;

/**
 *
 * @author arsalan
 */
public class ExcelWriter {

    //ExcelFile
    private File file;

    //WorkBook
    private WritableWorkbook book;

    //Writable Sheet
    private WritableSheet sheet;

    //private
    private int noOfSheets = 0;

    //file open
    private boolean isOpen = false;

    /**
     *
     * @param fileName
     */
    public ExcelWriter(String fileName) {
        try {
            file = new File(fileName);
            book = Workbook.createWorkbook(file, Workbook.getWorkbook(file));
            isOpen = true;
        } catch (Exception w) {
            createFile(fileName);
        }
    }

    /**
     * Create an excel file
     *
     * @param fileName
     * @return
     */
    public boolean createFile(String fileName) {
        try {
            file = new File(fileName);
            book = Workbook.createWorkbook(file);
            isOpen = true;
            return true;
        } catch (IOException ex) {
            Logger.getLogger(ExcelWriter.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
    }

    /**
     * Create new sheet
     *
     * @param sheetName
     * @return
     */
    public boolean createSheet(String sheetName) {
        if (isOpen) {
            sheet = book.createSheet(sheetName, (book.getNumberOfSheets() + 1));
            noOfSheets++;
            setSheet((book.getNumberOfSheets() - 1));
            return true;
        } else {
            System.out.println("File Not Open");
            return false;
        }
    }

    /**
     * Insert column value as string
     *
     * @param row
     * @param col
     * @param content
     * @return
     */
    public boolean writeString(int row, int col, String content) {
        try {
            if (isOpen) {
                Label label = new Label(col, row, content);
                sheet.addCell(label);

                return true;
            } else {
                System.out.println("File Not Open");
                return false;
            }

        } catch (WriteException ex) {
            Logger.getLogger(ExcelWriter.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
    }

    /**
     * Insert column value as number
     *
     * @param row
     * @param col
     * @param content
     * @return
     */
    public boolean writeNumber(int row, int col, double content) {
        try {
            if (isOpen) {
                Number num = new Number(col, row, content);
                sheet.addCell(num);

                return true;
            } else {
                System.out.println("File Not Open");
                return false;
            }
        } catch (WriteException ex) {
            Logger.getLogger(ExcelWriter.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
    }

    /**
     * close the file
     *
     * @return
     */
    public boolean closeFile() {
        try {
            if (isOpen) {
                isOpen = false;
                book.write();
                book.close();
                return true;
            } else {
                System.out.println("File Already Closed");
                return false;
            }

        } catch (IOException ex) {
            Logger.getLogger(ExcelWriter.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        } catch (WriteException ex) {
            Logger.getLogger(ExcelWriter.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
    }

    /**
     * Get total number of sheets
     *
     * @return
     */
    public int getNoOfSheets() {
        if (isOpen) {
            return noOfSheets;
        } else {
            System.out.println("File Not open");
            return -1;
        }
    }

    /**
     * Get all sheet names
     *
     * @return
     */
    public String[] sheetNames() {
        String[] str = null;

        if (isOpen) {
            str = book.getSheetNames();
        }

        return str;
    }

    /**
     * Set sheet number
     *
     * @param sheetIndex
     * @return
     */
    public boolean setSheet(int sheetIndex) {
        boolean check = false;
        if (isOpen) {
            if (sheetIndex < book.getNumberOfSheets()) {
                this.sheet = book.getSheet(sheetIndex);
                check = true;
            } else {
                System.out.println("Wrong Index");
            }
        } else {
            System.out.println("Book is close");
        }

        return check;
    }

}
