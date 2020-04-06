package com.excel2;

import com.excel2.Entity.Book;
import com.excel2.ReadExcel.readExcel;
import com.excel2.controller.controller;
import com.excel2.service.getBook;
import com.excel2.WriteExcel.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
import java.util.List;

@SpringBootApplication
public class DemoApplication {

    public static void main(String[] args) throws IOException {

        SpringApplication.run(DemoApplication.class, args);
//        getBook getBook = new getBook();
//        final List<Book> books = getBook.getBook();
//        final String excelFile = "E:\\demo\\listProduct.xls";
//        WriteExcelExample.writeExcel(books,excelFile);
        controller controller = new controller();
        int a = 922;
        String result = controller.convertLessThanOneThousand(a);
        System.out.println(result);
//
//        final List<Book> books1 = readExcel.readExcel(excelFile);
//         for (Book book: books){
//             System.out.print("----"+book.getId()+ "----");
//             System.out.println();
//             System.out.print(book.getTitle());
//             System.out.println();
//             System.out.print(book.getQuantity());
//             System.out.println();
//             System.out.print(book.getPrice());
//             System.out.println();
//             System.out.print(book.getPrice()*book.getQuantity());
//
//         }

    }

}
