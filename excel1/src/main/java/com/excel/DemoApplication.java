package com.excel;

import com.excel.model.Product;
import com.excel.service.ProductService;
import org.apache.poi.ss.util.CellReference;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import java.io.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.core.io.ByteArrayResource;

@SpringBootApplication
public class DemoApplication {



    public static void main(String[] args) {
        SpringApplication.run(DemoApplication.class, args);

         try {
             //khởi tạo excel
            HSSFWorkbook workbook = new HSSFWorkbook();
            //khởi tạo sheet(trang)
            HSSFSheet sheet = workbook.createSheet("List Product");
            //Tạo dòng thứ 0
            Row rowHeading = sheet.createRow(0);
            rowHeading.createCell(0).setCellValue("id");
            rowHeading.createCell(1).setCellValue("Name");
            rowHeading.createCell(2).setCellValue("Creation Date");
            rowHeading.createCell(3).setCellValue("Price");
            rowHeading.createCell(4).setCellValue("Quanlity");
            rowHeading.createCell(5).setCellValue("Sub Total");
            for (int i = 0;i<6;i++){
                CellStyle stylerowHeading = workbook.createCellStyle();

                //thiết lập định dạng
                Font font = workbook.createFont();
                font.setBold(true); // in đậm
                font.setFontName(HSSFFont.FONT_ARIAL);//font chữ
                font.setFontHeightInPoints((short) 11);//font size
                font.setColor(IndexedColors.WHITE.getIndex());//text color

                //cellStyle áp dụng font ở trên
                stylerowHeading.setFont(font);
                stylerowHeading.setFillForegroundColor(IndexedColors.BLUE.getIndex());//màu nền cho cell
                stylerowHeading.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                stylerowHeading.setBorderBottom(BorderStyle.THIN);

                //áp dụng định dạng trên cho 1 cell trên 1 dòng
                rowHeading.getCell(i).setCellStyle(stylerowHeading);
            }

            //đổ dữ liệu
            int r=1;
            ProductService productService = new ProductService();
            for (Product p : productService.findAll())
            {
                //tạo dòng thứ 2
                Row row = sheet.createRow(r);
                //khởi tạo ô id ở dòng r
                Cell cellId = row.createCell(0);
                //thiết lập giá trị cho ô đó
                cellId.setCellValue(p.getId());

                //khởi tạo ô name ở dong r
                Cell cellName = row.createCell(1);
                //thiết lập giá trị cho ô đó
                cellName.setCellValue(p.getName());

                //khởi tạo ô date
                Cell cellCreationDate =row.createCell(2);
                //thiết lập giá trị cho ô đó
                cellCreationDate.setCellValue(p.getCreationDate());

                CellStyle styleCreationDate = workbook.createCellStyle();
                //convert date
                HSSFDataFormat dfCreationDate = workbook.createDataFormat();
                styleCreationDate.setDataFormat(dfCreationDate.getFormat("m/d/yy"));
                //thiết lập convert date cho ô creationDate
                cellCreationDate.setCellStyle(styleCreationDate);

                // giá
                Cell cellPrice = row.createCell(3);
                cellPrice.setCellValue(p.getPrice());

                CellStyle stylePrice = workbook.createCellStyle();
                HSSFDataFormat cfPrice = workbook.createDataFormat();
                stylePrice.setDataFormat(cfPrice.getFormat("$#,##0.00"));
                cellPrice.setCellStyle(stylePrice);

                //số lượng
                Cell cellQuanlity = row.createCell(4);
                cellQuanlity.setCellValue(p.getQuanlity());


                //cột tổng tiền
                Cell cellTotal = row.createCell(5);
                cellTotal.setCellValue(p.getQuanlity()*p.getPrice());

                CellStyle styleTotal = workbook.createCellStyle();
                HSSFDataFormat cfTotal = workbook.createDataFormat();
                styleTotal.setDataFormat(cfTotal.getFormat("$#,##0.00"));
                cellTotal.setCellStyle(styleTotal);


                r++;
            }

            //Gộp column
             //dòng thứ 6
             Row rowTotal = sheet.createRow(productService.findAll().size()+1);
            Cell cellTextTotal = rowTotal.createCell(0);
            cellTextTotal.setCellValue("Total");
            //Gộp cột
            CellRangeAddress region = CellRangeAddress.valueOf("A7:E7");
            sheet.addMergedRegion(region);

            CellStyle styleTotal = workbook.createCellStyle();
            Font fontTextTotal = workbook.createFont();
            fontTextTotal.setBold(true);
            fontTextTotal.setFontHeightInPoints((short)11);
            fontTextTotal.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            styleTotal.setFont(fontTextTotal);
            cellTextTotal.setCellStyle(styleTotal);

            //Cột tổng của tổng tiền
             Cell cellTotalValue = rowTotal.createCell(5);
             cellTotalValue.setCellFormula("sum(F2:F6)");
             HSSFDataFormat dataFormat = workbook.createDataFormat();
             CellStyle styleTotalValue = workbook.createCellStyle();
             styleTotalValue.setDataFormat(dataFormat.getFormat("$#,##0.00"));
             cellTotalValue.setCellStyle(styleTotalValue);

            // tự dộng căn dãn nội dung trong 1 cell
             for (int i=0;i<6;i++){
                 sheet.autoSizeColumn(i);
             }



            //lưu file excel
             FileOutputStream out = new FileOutputStream(new File("E:\\demo\\listProduct.xls"));
             workbook.write(out);
             out.close();
             workbook.close();
             out.flush();
             System.out.println("Excel written successfully");


             ///////////////////////////////////////////////////
             //Example print ByteArrayResource
             ByteArrayResource resource = null;
             String pathExcelFile = null;
             try {
                 Path path = Paths.get("E:\\demo\\listProduct.xls");
                 byte[] bData = Files.readAllBytes(path);
                 resource = new ByteArrayResource(bData);

                 File excelFile = new File("listProduct.xls");
                 //lấy ra đường dẫn tuyệt đối của file
                  pathExcelFile = excelFile.getAbsolutePath();
                 if (excelFile.exists()){
                     excelFile.delete();
                 }
             }catch (IOException e){
                 e.printStackTrace();
             }
             System.out.println("Resource"+resource);

             //in đường dẫn tuyệt đối
             System.out.println("pathExcelFile"+ pathExcelFile);
         }catch (Exception e){
             System.out.println(e.getMessage());
         }
    }

}
