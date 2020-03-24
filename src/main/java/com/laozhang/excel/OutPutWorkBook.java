package com.laozhang.excel;

import java.util.List;
import java.util.Objects;

public abstract class OutPutWorkBook<T>{
    public final static String HEADER_TEMP_CLASS_PATH = "template/head";
    public final static String BODY_TEMP_CLASS_PATH = "template/body";
    public final static String DEFAULT_ENCODE = "UTF-8";

    public final static int SHEET_CAPACITY = 60000;

    protected String headerTemplate;
    protected String bodyTemplate;
    protected T headerRow;
    protected MyWorkSheet<T> sheet;
    private String sheetName;
    private int sheetIndex = 1;

    public OutPutWorkBook(){
        this(HEADER_TEMP_CLASS_PATH,BODY_TEMP_CLASS_PATH);
    }
    /**
     * @param bodyTemplate Sheet模板
     */
    public OutPutWorkBook(String bodyTemplate){
        this(HEADER_TEMP_CLASS_PATH,bodyTemplate);
    }
    /**
     * @param headerTemplate 文件头模板
     * @param bodyTemplate Sheet模板
     */
    public OutPutWorkBook(String headerTemplate,String bodyTemplate){
        this.headerTemplate = headerTemplate;
        this.bodyTemplate = bodyTemplate;
    }
    /**
     * 新建一个sheet
     * 同一个excel中每个sheet名不能重复
     * @param sheetName sheet名
     */
    public OutPutWorkBook<T> sheet(String sheetName){
        this.sheet = new MyWorkSheet<>();
        this.sheet.setName(sheetName);
        this.sheetName = sheetName;
        return this;
    }
    /**
     * 设置抬头
     * @param header 抬头行内容
     */
    public OutPutWorkBook<T> setHeader(T header){
        this.headerRow = header;
        this.sheet.setHeader(this.headerRow);
        return this;
    }
    /**
     * 填充数据
     * 可多次操作,超过60000自动转Sheet
     * @param data 批量数据
     */
    public OutPutWorkBook<T> fillSheet(List<T> data)throws Exception{
        data = Objects.requireNonNull(data);
        if(this.sheet == null){
            throw new RuntimeException("先新建一个sheet");
        }
        int rowCount = data.size();
        if(this.sheet.getRows() == null){
            if(rowCount > SHEET_CAPACITY){
                this.sheet.setRows(data.subList(0, SHEET_CAPACITY));
                write();
                this.sheet = new MyWorkSheet<>();
                this.sheet.setHeader(this.headerRow);
                this.sheet.setName(this.sheetName + "-" + this.sheetIndex ++);
                fillSheet(data.subList(SHEET_CAPACITY,rowCount));
            }else{
                this.sheet.setRows(data);
            }
        }else{
            int contentCount = this.sheet.getRows().size();
            if(rowCount + contentCount > SHEET_CAPACITY){
                int left = SHEET_CAPACITY - contentCount;
                this.sheet.addRows(data.subList(0, left));
                write();
                this.sheet = new MyWorkSheet<>();
                this.sheet.setHeader(this.headerRow);
                this.sheet.setName(this.sheetName + "-" + this.sheetIndex ++);
                fillSheet(data.subList(left,rowCount));
            }else{
                this.sheet.addRows(data);
            }
        }
        return this;
    }

    /**
     * 收尾
     */
    public void finish() throws Exception{
        if(this.sheet != null && this.sheet.getRows() != null && this.sheet.getRowNum() > 0){
            write();
        }
    }

    protected abstract void write() throws Exception;
}
