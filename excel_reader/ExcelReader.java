package excel_reader;

import java.io.File;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.*;

/**
 * Class to Read contents or find any content in an excel book
 * @author arsalan
 */
public class ExcelReader {

    /**
     * Object of Workbook to access Excel book
     */
    private Workbook book;
    /**
     * Object of Sheet to access specific sheet of excel book
     */
    private Sheet sheet;
    /**
     * Object of Cell to access specific cell of excel book
     */
    private Cell cell;
    /**
     * To store total number of sheets in a excel book
     */
    private int noOfSheets = 0;
    /**
     * To store maximum Row count of a specific sheet
     */
    private int maxRow = 0;
    /**
     * To store maximum Column count of a specific sheet
     */
    private int maxColumn = 0;
    /**
     * To check weather the file is open or not
     */
    private boolean isOpen = false;

    /**
     * Open an excel book with given link on creating object of this class.
     */
    public ExcelReader() {
    }

    /**
     * Open an excel book with given link.
     *
     * @param link
     * @return status weather the file extension is correct or not
     */
    public boolean openBook(String link) {
        boolean status = open(link);
        return status;
    }

    /**
     * close current excel book. Open an excel book with given link.
     *
     * @return status weather the file extension is correct or not. If extension
     * is correct then open this book. Stores total number of sheets in this
     * book. Set by default sheet to 0 and cell to 0,0
     */
    private boolean open(String link) {
        boolean status = false;
        if (link.substring(link.length() - 4, link.length()).equals(".xls")) {
            try {
                try {
                    if (isOpen) {
                        book.close();
                    }
                    book = Workbook.getWorkbook(new File(link));
                    isOpen = true;
                    noOfSheets = book.getNumberOfSheets();
                    setSheet(0);
                    cell = sheet.getCell(0, 0);
                    status = true;
                } catch (Exception e) {
                    isOpen = false;
                    System.out.println(e);
                }
            } catch (Exception exception) {
                isOpen = false;
                status = false;
                Logger.getLogger(ExcelReader.class.getName()).log(Level.SEVERE, null, exception);
            }
        } else {
            status = false;
            System.out.println("File Format is incorrect!");
        }

        return status;
    }

    /**
     * close the current book
     */
    public void closeBook() {
        if (isOpen) {
            book.close();
            isOpen = false;
        }
    }

    /**
     * @return Sheet names contained in an excel book if book is open. else
     * return null
     */
    public String[] sheetNames() {
        String[] str = null;

        if (isOpen) {
            str = book.getSheetNames();
        }

        return str;
    }

    /**
     * @return total number of sheets in an excel book if book is open. else
     * return 0.
     */
    public int getNoOfSheets() {
        int count = 0;

        if (isOpen) {
            count = noOfSheets;
        }

        return count;
    }

    public int getTotalRows() {
        return sheet.getRows();
    }

    public int getTotalCols() {
        return sheet.getColumns();
    }

    /**
     * Set the sheet to desired index.
     *
     * @param sheetIndex
     * @return weather the sheet is set or not(due to wrong sheet number).
     */
    public boolean setSheet(int sheetIndex) {
        boolean check = false;
        if (isOpen) {
            if (sheetIndex < noOfSheets) {
                this.sheet = book.getSheet(sheetIndex);
                maxRow = sheet.getRows();
                maxColumn = sheet.getColumns();
                check = true;
            } else {
                System.out.println("Wrong Index");
            }
        } else {
            System.out.println("Book is close");
        }

        return check;
    }

    /**
     * set cell of the sheet to desired row and column.
     *
     * @param row
     * @param column
     * @return weather the cell is correct or not.
     */
    public boolean setCell(int row, int column) {
        boolean check = false;
        if (isOpen) {
            if (row < maxRow && isOpen) {
                if (column < maxColumn) {
                    this.cell = sheet.getCell(column, row);
                    check = true;
                } else {
                    System.out.println("Wrong Column");
                }
            } else {
                System.out.println("Wrong Row");
            }
        } else {
            System.out.println("Book is close");
        }

        return check;
    }

    /**
     * Find key in an cell.
     *
     * @param key
     * @return weather the key is found or not.
     */
    public boolean exists(String key) {
        Cell cell;
        int row = 0, column = 0;
        boolean found = false;
        if (isOpen) {
            if (sheet.getRows() == 0 && sheet.getColumns() == 0) {
                found = false;
            } else {
                do {
                    column = 0;
                    do {
                        cell = sheet.getCell(column, row);
                        column++;
                        if (cell.getContents().equals(key)) {
                            found = true;
                        }
                    } while (column < sheet.getColumns() && !found);
                    row++;
                } while (!found && row < sheet.getRows());

                column--;
                row--;
            }
        } else {
            System.out.println("Book is close");
        }

        return found;
    }

    /**
     * Get row number of the cell which contains specific key.
     *
     * @param key
     * @return Integer which tells the row number 
     * If Key is not found so return -1 as row
     */
    public int getRow(String key) {
        int row = -1;
        try{
            row = sheet.findCell(key).getRow();
        }catch(Exception e){
            
        }
        return row;
    }

     /**
     * Get Column number of the cell which contains specific key.
     *
     * @param key
     * @return Integer which tells the column number 
     * If Key is not found so return -1 as row
     */
    public int getColumn(String key) {
        int column = -1;
        try{
            column = sheet.findCell(key).getColumn();
        }catch(Exception e){
            
        }
        return column;
    }
    
    /**
     * Get coordinates of the cell which contains specific key.
     *
     * @param key
     * @return ArrayList which contains row(at 0th location) and column(at 1st
     * location) of ArrayList. if Key is not found so return -1 as row and
     * column.
     */
    public ArrayList<Integer> getLocation(String key) {
        ArrayList<Integer> list = new ArrayList<>();
        try {
            Cell cell;
            int row = 0, column = 0;
            boolean found = false;
            if (isOpen) {
                if (sheet.getColumns() == 0 && sheet.getRows() == 0) {
                    found = false;
                } else {
                    do {
                        column = 0;
                        do {
                            cell = sheet.getCell(column, row);
                            column++;
                            if (cell.getContents().equals(key)) {
                                found = true;
                            }
                        } while (column < sheet.getColumns() && !found);
                        row++;
                    } while (!found && row < sheet.getRows());

                    column--;
                    row--;
                }
            } else {
                System.out.println("Book is close");
            }
            if (!found) {
                list.add(-1);
                list.add(-1);
            } else {
                list.add(row);
                list.add(column);
            }
        } catch (Exception exception) {
        }
        return list;
    }

    /**
     * @param row
     * @param column
     * @return content at specific row and column if these coordinates exist
     * else return "".
     */
    public String getContents(int row, int column) {
        String contents = "";
        if (isOpen) {
            if (row < maxRow && column < maxColumn) {
                Cell cell = sheet.getCell(column, row);
                contents = cell.getContents();
            }
        } else {
            System.out.println("Book is close");
        }
        return contents;
    }

    /**
     * Get Contents of selected Row if this row exist.
     *
     * @param row
     * @return ArrayList with contents in a row if the row exist, else return
     * empty ArrayList if row doesn't exist
     */
    public ArrayList<String> getContentsOfRow(int row) {
        ArrayList<String> contents = new ArrayList<>();

        if (isOpen) {
            if (row < maxRow) {
                Cell[] cells = sheet.getRow(row);
                for (Cell cell : cells) {
                    contents.add(cell.getContents());
                }
            }
        } else {
            System.out.println("Book is close");
        }
        return contents;
    }

    /**
     * Get Contents of selected Column if this row exist.
     *
     * @param column
     * @return ArrayList with contents in a column if the row exist, else return
     * empty ArrayList if column doesn't exist
     */
    public ArrayList<String> getContentsOfColumn(int column) {
        ArrayList<String> contents = new ArrayList<>();
        if (isOpen) {
            if (column < maxColumn) {
                Cell[] cells = sheet.getColumn(column);
                for (Cell cell : cells) {
                    contents.add(cell.getContents());
                }
            }
        } else {
            System.out.println("Book is close");
        }
        return contents;
    }

    /**
     * @return sheet name of current sheet
     */
    public String getSheetName() {
        if (isOpen) {
            return this.sheet.getName();
        } else {
            return "";
        }
    }

    /**
     * @return current row number of a cell
     * -1 if book is close
     */
    public int getRow() {
        int count = -1;
        if (isOpen) {
            count = cell.getRow();
        }

        return count;
    }

    /**
     * @return current column number of a cell
     *  -1 if book is close
     */
    public int getColumn() {
        int count = -1;
        if (isOpen) {
            count = cell.getColumn();
        }
        return count;
    }
    /**
     * @param args the command line arguments
     */
//    public static void main(String[] args) throws Exception {
//        // TODO code application logic here
//        ExcelReader read = new ExcelReader("test1.xls");
//        read.closeBook();
//        System.out.println(read.getLocation("ID"));
//    }
}
