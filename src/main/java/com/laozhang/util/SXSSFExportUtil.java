package com.laozhang.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

public class SXSSFExportUtil {
	public interface RowAction{
		void doAction(Row row, Map<String, Object> rowData);
	}

	private final static SimpleDateFormat todate = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
    private String[] header;
    private String[] headerMap2Body;
    private SXSSFWorkbook wb;
    private Sheet curSheet;
    private CellStyle headerStyle;
    private CellStyle bodyStyle;
    private int index;

    public SXSSFExportUtil(){
        this(5000);
    }
    public SXSSFExportUtil(int batch){
        this.wb = new SXSSFWorkbook(batch);
    }

    public SXSSFExportUtil setHeader(List<String> headCollection){
        headCollection = Objects.requireNonNull(headCollection);
        this.header = (String[])headCollection.toArray();
        return this;
    }

    public SXSSFExportUtil setHeader(String[] headCollection){
        this.header = Objects.requireNonNull(headCollection);
        return this;
    }

    public SXSSFExportUtil setHeaderStyle(CellStyle style){
        this.headerStyle = style;
        return this;
    }

    public SXSSFExportUtil setBodyStyle(CellStyle style){
        this.bodyStyle = style;
        return this;
    }

    public SXSSFExportUtil createSheet(){
        this.curSheet = this.wb.createSheet();
        initHeader();
        return this;
    }

    public SXSSFExportUtil createSheet(String sheetName){
        this.curSheet = this.wb.createSheet(sheetName);
        initHeader();
        return this;
    }

    public SXSSFExportUtil setMapping(String[] headerMap2Body){
        this.headerMap2Body = headerMap2Body;
        return this;
    }

    public SXSSFExportUtil fileIn(List<Map<String,Object>> rows){
        if(this.curSheet == null){
            throw new RuntimeException("need createSheet() first!");
        }
        if(rows != null && rows.size() > 0) {
            if(this.headerMap2Body == null){
                fillValue(rows);
            }else{
                fillValueWithMapping(rows);
            }
        }
        return this;
    }
    
    public SXSSFExportUtil fileIn(List<Map<String,Object>> rows, RowAction action){
        if(this.curSheet == null){
            throw new RuntimeException("need createSheet() first!");
        }
        if(rows != null && rows.size() > 0) {
        	Row rowFrame = null;
        	for(Map<String,Object> row : rows){
                rowFrame = this.curSheet.createRow(index++);
        		action.doAction(rowFrame, row);
        	}
        }
        return this;
    }

    public void write(OutputStream outputStream){
        outputStream = Objects.requireNonNull(outputStream);
        try{
            this.wb.write(outputStream);
            outputStream.flush();
        }catch(Exception ex){

        }finally {
            try {
                outputStream.close();
                this.wb.dispose();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private void initHeader(){
        Cell cell;
        int len;
        this.index = 0;
        if(this.header != null && (len = this.header.length) > 0) {
            Row row = this.curSheet.createRow(this.index++);
            for (int i = 0; i < len; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(this.headerStyle);
                cell.setCellValue(header[i]);
            }
        }
    }

    private void fillValue(List<Map<String,Object>> rows){
        Cell cell;
        Map<String,Object> rowMap;
        Iterator<Map<String,Object>> iterator = rows.iterator();
        while(iterator.hasNext()) {
            rowMap = iterator.next();
            Row row = this.curSheet.createRow(index++);
            int cellIndex = 0;
            for (Map.Entry<String,Object> item : rowMap.entrySet()) {
                cell = row.createCell(cellIndex++);
                cell.setCellStyle(this.bodyStyle);
                convertCell(cell,item.getValue());
            }
        }
    }

    private void fillValueWithMapping(List<Map<String,Object>> rows){
        Cell cell;
        Map<String,Object> cellMap;
        Iterator<Map<String,Object>> iterator = rows.iterator();
        while(iterator.hasNext()) {
            cellMap = iterator.next();
            Row row = this.curSheet.createRow(index++);
            for (int i = 0, cellLen = header.length; i < cellLen; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(this.bodyStyle);
                convertCell(cell,cellMap.get(this.headerMap2Body[i]));
            }
        }
    }

    private void convertCell(Cell cell,Object v){
    	String str;
        if (v instanceof String) {
            cell.setCellValue((String) v);
        } else if (v instanceof Number) {
            cell.setCellValue(Double.valueOf(v.toString()));
        } else if (v instanceof Date) {
        	str = todate.format(v);
            cell.setCellValue(str);
        } else if (v instanceof Boolean) {
            cell.setCellValue((boolean) v);
        }
    }
}
