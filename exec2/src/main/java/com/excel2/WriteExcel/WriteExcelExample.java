package com.excel2.WriteExcel;

import com.excel2.Entity.Book;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

public class WriteExcelExample {
    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_TITLE = 1;
    public static final int COLUMN_INDEX_PRICE = 2;
    public static final int COLUMN_INDEX_QUANTITY = 3;
    public static final int COLUMN_INDEX_TOTAL = 4;
    public static void writeExcel(List<Book> books, String excelFile) {
        // tạo workbook
        Workbook workbook = getWorkbook(excelFile);

        //tạo sheet
        Sheet sheet = workbook.createSheet("Books");

        //Tạo tiêu đề header
        int rowIndex = 0;
        writeHeader(sheet,rowIndex);

        //tự động căn chỉnh
        //lấy ra tổng số cột của dòng
        int numberOfColumn = sheet.getRow(0).getPhysicalNumberOfCells();
        autosizeColumn(sheet,numberOfColumn);

        //đổ dữ liệu
        rowIndex++;
        for (Book book : books) {
            // Create row
            Row row = sheet.createRow(rowIndex);
            // Write data on row
            writeBook(book, row);
            rowIndex++;
        }

        writeFooter(sheet,rowIndex);
        //lưu file
        createOutputFile(workbook,excelFile);
        System.out.println("Done");

    }

    //khởi tạo workbook
    private static Workbook getWorkbook(String excelFilePath){
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        }else if(excelFilePath.endsWith("xls")){
            workbook = new HSSFWorkbook();
        }else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
        return workbook;
    }

    private static void writeHeader(Sheet sheet,int rowIndex){
        CellStyle cellStyle = createStyleForHeader(sheet);

        Row row = sheet.createRow(rowIndex);

        Cell cell = row.createCell(0);
        cell.setCellValue("Id");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        cell.setCellValue("Title");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Price");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(3);
        cell.setCellValue("Quanlity");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(4);
        cell.setCellValue("Total money");
        cell.setCellStyle(cellStyle);
    }

    private static CellStyle createStyleForHeader(Sheet sheet){
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short)14);
        font.setColor(IndexedColors.WHITE.getIndex());

        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        return  cellStyle;
    }

    private static void createOutputFile(Workbook workbook, String excelFile){
        try (OutputStream os = new FileOutputStream(excelFile)){
            workbook.write(os);
        }catch (Exception e){
            e.getMessage();
        }
    }
    //tự động căn chỉnh
    private static void autosizeColumn(Sheet sheet,int lastColumn){
        for (int columnIndex = 0;columnIndex<lastColumn;columnIndex++){
            sheet.autoSizeColumn(columnIndex);
        }
    }


    private static CellStyle cellStyleFormatNumber = null;
    //Đổ dữ liệu
    private static void writeBook(Book book,Row row){
        if (cellStyleFormatNumber == null) {
            // Format number
            short format = (short)BuiltinFormats.getBuiltinFormat("#,##0");
            // DataFormat df = workbook.createDataFormat();
            // short format = df.getFormat("#,##0");

            //Create CellStyle
            Workbook workbook = row.getSheet().getWorkbook();
            cellStyleFormatNumber = workbook.createCellStyle();
            cellStyleFormatNumber.setDataFormat(format);
        }

        Cell cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellValue(book.getId());

        cell = row.createCell(COLUMN_INDEX_TITLE);
        cell.setCellValue(book.getTitle());

        cell = row.createCell(COLUMN_INDEX_PRICE);
        cell.setCellValue(book.getPrice());
        cell.setCellStyle(cellStyleFormatNumber);

        cell = row.createCell(COLUMN_INDEX_QUANTITY);
        cell.setCellValue(book.getQuantity());

        // Create cell formula
        // totalMoney = price * quantity
        cell = row.createCell(COLUMN_INDEX_TOTAL, CellType.FORMULA);
        cell.setCellStyle(cellStyleFormatNumber);
        int currentRow = row.getRowNum() + 1;
        String columnPrice = CellReference.convertNumToColString(COLUMN_INDEX_PRICE);
        String columnQuantity = CellReference.convertNumToColString(COLUMN_INDEX_QUANTITY);
        cell.setCellFormula(columnPrice + currentRow + "*" + columnQuantity + currentRow);

    }

    private static void writeFooter(Sheet sheet,int rowindex){
        Row row = sheet.createRow(rowindex);
        Cell cell = row.createCell(COLUMN_INDEX_TOTAL,CellType.FORMULA);
        cell.setCellFormula("SUM(E2:E6)");
    }
}
