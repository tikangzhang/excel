package com.laozhang.excel;

import org.antlr.stringtemplate.StringTemplate;
import org.antlr.stringtemplate.StringTemplateGroup;

import java.io.*;

public class FileWorkBook<T> extends OutPutWorkBook<T>{
    private String fileName;
    /**
     * @param fileName 欲生成或追加内容的文件路径
     */
    public FileWorkBook(String fileName){
        this(fileName,HEADER_TEMP_CLASS_PATH,BODY_TEMP_CLASS_PATH);
    }
    /**
     * @param fileName 欲生成或追加内容的文件路径
     * @param bodyTemplate Sheet模板
     */
    public FileWorkBook(String fileName,String bodyTemplate){
        this(fileName,HEADER_TEMP_CLASS_PATH,bodyTemplate);
    }
    /**
     * @param fileName 欲生成或追加内容的文件路径
     * @param headerTemplate 文件头模板
     * @param bodyTemplate Sheet模板
     */
    public FileWorkBook(String fileName,String headerTemplate,String bodyTemplate){
        super(headerTemplate,bodyTemplate);
        this.fileName = fileName;
    }

    @Override
    protected void write() throws Exception {
        File file = new File(this.fileName);
        if(!file.exists()){
            file.createNewFile();
            generate(file);
        }else {
            append(file);
        }
    }

    private void append(File file) throws Exception {
        StringTemplateGroup stGroup = new StringTemplateGroup("stringTemplate");
        RandomAccessFile writer = null;
        try {
            writer =  new RandomAccessFile(file,"rw");
            writer.seek(writer.length() - 11);
            StringTemplate body =  stGroup.getInstanceOf(this.bodyTemplate);
            body.setAttribute("worksheet", this.sheet);
            writer.write(body.toString().getBytes(DEFAULT_ENCODE));
            writer.write("</Workbook>".getBytes(DEFAULT_ENCODE));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(writer != null){
                try {
                    writer.close();
                } catch (IOException e) {
                }
            }
        }
    }

    private void generate(File file) {
        StringTemplateGroup stGroup = new StringTemplateGroup("stringTemplate");
        StringTemplate head =  stGroup.getInstanceOf(this.headerTemplate);
        BufferedOutputStream writer = null;
        try {
            writer = new BufferedOutputStream(new FileOutputStream(file));
            writer.write(head.toString().getBytes(DEFAULT_ENCODE));
            writer.flush();
            StringTemplate body =  stGroup.getInstanceOf(this.bodyTemplate);
            body.setAttribute("worksheet", this.sheet);
            writer.write(body.toString().getBytes(DEFAULT_ENCODE));
            writer.write("</Workbook>".getBytes(DEFAULT_ENCODE));
            writer.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(writer != null){
                try {
                    writer.close();
                } catch (IOException e) {
                }
            }
        }
    }
}

