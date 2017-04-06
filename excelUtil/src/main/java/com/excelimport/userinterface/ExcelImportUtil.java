package com.excelimport.userinterface;

import com.excelimport.bean.ExcelData;
import com.excelimport.bean.ExcelStruct;
import com.excelimport.exception.ExcelImportException;
import com.excelimport.util.ExcelDataReader;
import com.excelimport.util.ParseXMLUtil;
import org.apache.commons.lang3.StringUtils;
import org.dom4j.Document;
import org.dom4j.io.SAXReader;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * 读取导入的Excel的内容 模板要求： （1）开始重复行与End行 有且只能有 一空行
 * Created by can on 2017/3/28.
 */
public class ExcelImportUtil {
    private ExcelImportUtil(){}

    /**
     * 读取Excel文件
     * @param xmlFilePath 字段field与单元格cellRef映射xml文件
     * @param excelFilePath 需要读取的excel文件路径
     * @return
     */
    public static ExcelData readExcel(String xmlFilePath, String excelFilePath) throws ExcelImportException{
        if(StringUtils.isEmpty(xmlFilePath) || StringUtils.isEmpty(excelFilePath)){
            return null;
        }

        try {
            // 1. 解析XML描述文件
            ExcelStruct excelStruct = ParseXMLUtil.parseImportStruct(xmlFilePath);
            // 2. 按照XML描述文件，来解析Excel中文件的内容
            return ExcelDataReader.readExcel(excelStruct, excelFilePath, 0);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new ExcelImportException("导入Excel失败 - XML描述文件未找到 : ", e);
        } catch (IOException e) {
            e.printStackTrace();
            // log.error("导入Excel失败 - IO异常 : ", e);
            throw new ExcelImportException("导入Excel失败 - IO异常 : ", e);
        } catch (Exception e) {
            e.printStackTrace();
            // log.error("导入Excel失败 : ", e);
            throw new ExcelImportException("导入Excel失败 : ", e);
        }
    }
}
