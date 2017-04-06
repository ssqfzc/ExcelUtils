package com.excelimport;

        import com.excelimport.bean.ExcelData;
        import com.excelimport.userinterface.ExcelImportUtil;

/**
 * Created by can on 2017/3/21.
 */
public class Test {
    private static final String xmlFile = "D:\\excel_desc.xml";
    private static final String excelFile = "D:\\info_CRM.xls";

    public static void main(String[] args){
        ExcelData data = ExcelImportUtil.readExcel(xmlFile, excelFile);
        System.out.println(data);
    }
}
