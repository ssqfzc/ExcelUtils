package com.excelimport.util;

import com.excelimport.bean.ExcelData;
import com.excelimport.bean.ExcelStruct;
import org.apache.commons.lang3.StringUtils;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

import static org.junit.Assert.*;

/**
 * Created by can on 2017/4/2.
 */
public class ParseXMLUtilTest {
    private static final String xmlFile = "D:\\excel_desc.xml";
    private List onceList;
    private ExcelStruct excelStruct;

    @Before
    public void setUp() throws Exception {
        InputStream is = new FileInputStream(new File(xmlFile));
        if (is == null) {
            throw new FileNotFoundException("Excel的描述文件 : " + xmlFile + " 未找到.");
        }
        SAXReader saxReader = new SAXReader();
        Document document = saxReader.read(is);
        // 根节点
        Element root = document.getRootElement();
        // 一次导入
        onceList = root.elements("onceImport");
        // 重复导入
        List repeatList = root.elements("repeatImport");

        // 校验器的定义
        List validators = root.elements("validators");
        // 单元格校验
        List cellValidators = root.elements("cell-validators");

        excelStruct = new ExcelStruct();

        // 读取校验器配置
//        parseValidatorConfig(excelStruct, validators, cellValidators);

//        simpleParseOnceImport(excelStruct, onceList);

        is.close();

    }

    @Test
    public void parseImportStruct() throws Exception {
        System.out.println(ParseXMLUtil.parseImportStruct(xmlFile));
    }

    @Test
    public void simpleParseOnceImport() throws Exception{
        System.out.println(("=sdf".split("=")).length);
    }


}