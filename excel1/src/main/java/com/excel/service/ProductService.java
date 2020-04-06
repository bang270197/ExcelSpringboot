package com.excel.service;

import com.excel.model.Product;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Service
public class ProductService {

    public List<Product> findAll(){
        try{
            List<Product> result = new ArrayList<Product>();
            for (int i =1;i<6;i++){
                result.add(new Product("id"+i,"bang "+i,i*1000,i+1,new Date()));
            }
            return result;
        }catch (Exception e){
            return null;
        }
    }
}
