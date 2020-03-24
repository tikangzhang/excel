package com.laozhang.util;

import java.io.*;

public class UsingSXSSFExportExcel {
    static String outPutFileName = "temp/output.xls";

    public static void forAny() throws Exception {
        SXSSFExportUtil utilObj = new SXSSFExportUtil();
        utilObj
                .setHeader(new String[]{"时间","地点","姓甚","名谁","雌雄","主要联系人","次要联系人","做啥了？"})
                .createSheet("sheet1")
                .fileIn(UsingNewWayExportExcel.buildMapData(100000))
                .write(new FileOutputStream(new File(outPutFileName)));
    }
}
