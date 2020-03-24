package com.laozhang.excel;

import org.antlr.stringtemplate.StringTemplate;
import org.antlr.stringtemplate.StringTemplateGroup;

import java.io.BufferedOutputStream;
import java.io.OutputStream;

public class StreamWorkBook<T> extends OutPutWorkBook<T>{
    private OutputStream outputStream;
    private boolean started;
    /**
     * @param outputStream 输出流，不支持追加
     */
    public StreamWorkBook(OutputStream outputStream){
        this(outputStream,HEADER_TEMP_CLASS_PATH,BODY_TEMP_CLASS_PATH);
    }
    /**
     * @param outputStream 输出流，不支持追加
     * @param bodyTemplate Sheet模板
     */
    public StreamWorkBook(OutputStream outputStream,String bodyTemplate){
        this(outputStream,HEADER_TEMP_CLASS_PATH,bodyTemplate);
    }
    /**
     * @param outputStream 输出流，不支持追加
     * @param headerTemplate 文件头模板
     * @param bodyTemplate Sheet模板
     */
    public StreamWorkBook(OutputStream outputStream,String headerTemplate,String bodyTemplate){
        super(headerTemplate,bodyTemplate);
        this.outputStream = new BufferedOutputStream(outputStream);
    }

    @Override
    protected void write() throws Exception {
        StringTemplateGroup stGroup = new StringTemplateGroup("stringTemplate");
        if(!this.started){
            StringTemplate head = stGroup.getInstanceOf(this.headerTemplate);
            this.outputStream.write(head.toString().getBytes(DEFAULT_ENCODE));
            this.outputStream.flush();
            this.started = true;
        }
        StringTemplate body =  stGroup.getInstanceOf(this.bodyTemplate);
        body.setAttribute("worksheet", this.sheet);
        this.outputStream.write(body.toString().getBytes(DEFAULT_ENCODE));
    }
    /**
     * 收尾
     */
    public void finish() throws Exception{
        if(this.sheet != null && this.sheet.getRows() != null && this.sheet.getRowNum() > 0){
            write();
        }
        this.outputStream.write("</Workbook>".getBytes(DEFAULT_ENCODE));
        this.outputStream.flush();
        this.outputStream.close();
    }
}
