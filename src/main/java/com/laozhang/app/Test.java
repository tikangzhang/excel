package com.laozhang.app;


import com.laozhang.util.UsingNewWayExportExcel;
import com.laozhang.util.UsingSXSSFExportExcel;

public class Test {
    public static void main(String[] args) throws Exception{
        long cur = System.currentTimeMillis();
        //UsingNewWayExportExcel.forObject();
        UsingNewWayExportExcel.forMap();
        //UsingNewWayExportExcel.forStreamObject();
        //UsingSXSSFExportExcel.forAny();
        System.out.println(String.format("耗时：%d毫秒",System.currentTimeMillis() - cur));
    }
}
