package com.excel2.ReadExcel;

import com.excel2.Entity.Book;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class readExcel {
    public static List<Book> readExcel(String excelFilePath) throws IOException {
        List<Book> bookList = new ArrayList<>();

        //lấy file
        InputStream inputStream = new FileInputStream(new File(excelFilePath));

        //lấy workbook
        Workbook workbook = getWork(inputStream,excelFilePath);

        //lấy sheet
        Sheet sheet = workbook.getSheetAt(0);

        //lấy row từ sheet
        Iterator<Row> iterable = sheet.iterator();
        while (iterable.hasNext()){
            Row nextRow = iterable.next();
            if (nextRow.getRowNum()==0){
                continue;
            }

            //lấy cell
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            //đọc các giá trị trong cell và set giá trị cho book
            Book book = new Book();
            while (cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                Object cellValue = getCellValue(cell);
                if (cellValue == null || cellValue.toString().isEmpty()) {
                    continue;
                }
                //thiết lập giá trị cho book
                int columnIndex = cell.getColumnIndex();
                switch (columnIndex)
                {
                    case 0:
                        book.setId(new BigDecimal((double)cellValue).intValue());
                        break;
                    case 1:
                        book.setTitle((String)getCellValue(cell));
                        break;
                    case 2:
                        book.setQuantity(new BigDecimal((double)cellValue).intValue());
                        break;
                    case 3:
                        book.setPrice((double)getCellValue(cell));
                        break;
                    case 4:
                        book.setTotalMoney((double)getCellValue(cell));
                        break;
                    default:
                        break;
                }

            }
            bookList.add(book);
        }
        workbook.close();
        inputStream.close();

        return bookList;

    }

    //đổ ra workbook từ file
    private static Workbook getWork(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")){
            workbook = new XSSFWorkbook(inputStream);
        }else if (excelFilePath.endsWith("xls")){
            workbook = new HSSFWorkbook(inputStream);
        }else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
        return workbook;
    }

    private static Object getCellValue(Cell cell){
        CellType cellType = cell.getCellTypeEnum();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
            }
            return cellValue;
        }
    }

