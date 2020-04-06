package com.excel.model;

import java.util.Date;

public class Product {
    private String id;
    private String name;
    private float price;
    private int quanlity;
    private Date creationDate;

    public Product() {
    }

    public Product(String id, String name, float price, int quanlity, Date creationDate) {
        this.id = id;
        this.name = name;
        this.price = price;
        this.quanlity = quanlity;
        this.creationDate = creationDate;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public float getPrice() {
        return price;
    }

    public void setPrice(float price) {
        this.price = price;
    }

    public int getQuanlity() {
        return quanlity;
    }

    public void setQuanlity(int quanlity) {
        this.quanlity = quanlity;
    }

    public Date getCreationDate() {
        return creationDate;
    }

    public void setCreationDate(Date creationDate) {
        this.creationDate = creationDate;
    }
}
