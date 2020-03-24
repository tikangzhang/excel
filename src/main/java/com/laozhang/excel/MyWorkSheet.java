package com.laozhang.excel;

import java.util.LinkedList;
import java.util.List;

public class MyWorkSheet<T>{
    private String name;

    private int rowNum = 0;

    private T header;

    private List<T> rows;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public T getHeader() {
        return header;
    }

    public void setHeader(T header) {
        this.header = header;
    }

    public List<T> getRows() {
        return rows;
    }

    public void setRows(List<T> rows) {
        this.rows = new LinkedList<>(rows);
        reCalcSize();
    }

    public void addRows(List<T> data){
        this.rows.addAll(data);
        reCalcSize();
    }

    private void reCalcSize(){
        this.rowNum = this.rows.size() + 1;
    }
}
