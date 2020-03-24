package com.laozhang.util;

import com.laozhang.entity.Row;
import com.laozhang.excel.FileWorkBook;
import com.laozhang.excel.StreamWorkBook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class UsingNewWayExportExcel {
    static String outPutFileName = "temp/newOutput.xls";

    public static void forStreamObject()throws Exception{
        com.laozhang.entity.Row header = buildHeader();

        List<com.laozhang.entity.Row> data1 = buildData(10);
        List<com.laozhang.entity.Row> data2 = buildData(30000);
        List<com.laozhang.entity.Row> data3 = buildData(30000);

        FileOutputStream outputStream = new FileOutputStream(new File("temp/output.xls"));
        new StreamWorkBook<com.laozhang.entity.Row>(outputStream)
                .sheet("sheet1")
                .setHeader(header)
                .fillSheet(data1)
                .fillSheet(data2)
                .fillSheet(data3)
                .finish();
    }

    public static void forObject()throws Exception{
        com.laozhang.entity.Row header = buildHeader();

        List<com.laozhang.entity.Row> data1 = buildData(120000);
        List<com.laozhang.entity.Row> data2 = buildData(20000);
        List<com.laozhang.entity.Row> data3 = buildData(29999);
        List<com.laozhang.entity.Row> data4 = buildData(60000);
        List<com.laozhang.entity.Row> data5 = buildData(1);

        // output.xls存在就追加，不存在就新建
        new FileWorkBook<com.laozhang.entity.Row>(outPutFileName)
                .sheet("sheet1")
                .setHeader(header)
                .fillSheet(data1)
                .fillSheet(data2)
                .fillSheet(data3)
                .fillSheet(data4)
                .fillSheet(data5)
                .finish();
    }

    public static void forMap()throws Exception{
        Map<String, Object> header = buildMapHeader();
        List<Map<String,Object>> data1 = buildMapData(60001);
        List<Map<String,Object>> data2 = buildMapData(20000);
        List<Map<String,Object>> data3 = buildMapData(29999);
        List<Map<String,Object>> data4 = buildMapData(60000);
        List<Map<String,Object>> data5 = buildMapData(1);
        List<Map<String,Object>> data6 = buildMapData(360000);

        // output.xls存在就追加，不存在就新建
        FileWorkBook<Map<String,Object>> file = new FileWorkBook<Map<String,Object>>(outPutFileName);
        file.sheet("页1")
                .setHeader(header)
                .fillSheet(data1)
                .fillSheet(data2)
                .fillSheet(data3)
                .fillSheet(data4)
                .fillSheet(data5)
                .fillSheet(data6)
                .finish();
        file = new FileWorkBook<Map<String,Object>>(outPutFileName);
        file.sheet("页2")
                .setHeader(header)
                .fillSheet(data1)
//                .fillSheet(data2)
//                .fillSheet(data3)
//                .fillSheet(data4)
//                .fillSheet(data5)
//                .fillSheet(data6)
                .finish();
    }

    public static com.laozhang.entity.Row buildHeader(){
        com.laozhang.entity.Row header = new com.laozhang.entity.Row();
        header.setF0("时间");
        header.setF1("地点");
        header.setF2("姓甚");
        header.setF3("名谁");
        header.setF4("雌雄");
        header.setF5("主要联系人");
        header.setF6("次要联系人");
        header.setF7("做啥了？");
        return header;
    }
    public static List<com.laozhang.entity.Row> buildData(int howmuch){
        com.laozhang.entity.Row row;
        List<com.laozhang.entity.Row> data = new LinkedList<>();
        for(int j = 0;j < howmuch;j++){
            row = new Row();
            row.setF0("冰河世纪" + j);
            row.setF1("柏拉图包房" + (j % 12));
            row.setF2("李");
            row.setF3("老"+ (j % 8));
            row.setF4((j & 2) == 0 ? "男" : "女");
            row.setF5("隔壁老王" + j);
            row.setF6("后院王大爷" + j);
            row.setF7((j & 2) == 0 ? "杀人" : "放火");
            data.add(row);
        }
        return data;
    }

    public static Map<String,Object> buildMapHeader(){
        Map<String,Object> header = new HashMap<>();
        header.put("f0","时间");
        header.put("f1","地点");
        header.put("f2","姓甚");
        header.put("f3","名谁");
        header.put("f4","雌雄");
        header.put("f5","主要联系人");
        header.put("f6","次要联系人");
        header.put("f7","做啥了？");
        return header;
    }
    public static List<Map<String,Object>> buildMapData(int howmuch){
        Map<String,Object> row = null;
        List<Map<String,Object>> data = new LinkedList<>();
        for(int j=0;j < howmuch;j++){
            row = new HashMap<>();
            row.put("f0","冰河世纪" + j);
            row.put("f1","柏拉图包房" + (j % 12));
            row.put("f2","钱");
            row.put("f3","老"+ (j % 8));
            row.put("f4",(j & 2) == 0 ? "男" : "女");
            row.put("f5","隔壁老王" + j);
            row.put("f6","后院王大爷" + j);
            row.put("f7",(j & 2) == 0 ? "杀人" : "放火");
            data.add(row);
        }
        return data;
    }
}
