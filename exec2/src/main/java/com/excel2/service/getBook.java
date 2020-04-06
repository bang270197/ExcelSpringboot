package com.excel2.service;

import com.excel2.Entity.Book;

import java.util.ArrayList;
import java.util.List;

public class getBook {
    public List<Book> getBook(){
        List<Book> bookList = new ArrayList<>();
        for (int i = 1;i<6;i++){
           Book book = new Book(i,"Book"+i,i*2,i*1000.0);
           bookList.add(book);
        }
        return bookList;
    }
}
