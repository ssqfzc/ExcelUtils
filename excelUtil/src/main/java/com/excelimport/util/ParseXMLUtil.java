package com.excelimport.util;

import com.excelimport.bean.ExcelStruct;
import com.excelimport.bean.ImportCellDesc;
import com.excelimport.exception.ExcelImportException;
import org.apache.commons.lang3.StringUtils;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 解析导入Excel的XML描述文件
 */
@SuppressWarnings("unchecked")
public class ParseXMLUtil {
    private ParseXMLUtil() {}

    /**
     * 根据给定的XML文件解析出Excel的结构
     */
    public static ExcelStruct parseImportStruct(String xmlFile) throws DocumentException, IOException {
        if (StringUtils.isEmpty(xmlFile)) {
            return null;
        }
        InputStream is = new FileInputStream(new File(xmlFile));
        if (is == null) {
            throw new FileNotFoundException("Excel的描述文件 : " + xmlFile + " 未找到.");
        }
        SAXReader saxReader = new SAXReader();
        Document document = saxReader.read(is);
        // 根节点
        Element root = document.getRootElement();
        // 一次导入
        List onceList = root.elements("onceImport");
        // 重复导入
        List repeatList = root.elements("repeatImport");

        // 校验器的定义
        List validators = root.elements("validators");
        // 单元格校验
        List cellValidators = root.elements("cell-validators");

        ExcelStruct excelStruct = new ExcelStruct();

        // 读取校验器配置
        parseValidatorConfig(excelStruct, validators, cellValidators);
        // 获取需要一次性导入的xml解析文件
        simpleParseOnceImport(excelStruct, onceList);
        // 获取excel重复读取的xml解析文件
        simpleParseRepeatImport(excelStruct, repeatList);
        is.close();

        return excelStruct;
    }

    /**
     * 解析重复导入xml文档
     * @param excelStruct Excel导入描述文件的结构
     * @param repeatList 重复导入集合
     */
    private static void simpleParseRepeatImport(ExcelStruct excelStruct, List repeatList) {
        if (repeatList == null || repeatList.size() <= 0) {
            return;
        }
        List<ImportCellDesc> importCells = getImportCellDescList(excelStruct,repeatList);
        if (importCells == null || importCells.size() <= 0) {
            return;
        }
        // 读取终止符
        String endCode = null;
        Element repElem = (Element) repeatList.get(0);
        try {
            endCode = ((Element) repElem.elements("endCode").get(0)).getTextTrim();
        } catch (IndexOutOfBoundsException e) {
            throw new ExcelImportException("导入Excel失败 : 请在XML描述文件中添加<endCode/>配置项!");
        }
        excelStruct.setEndCode(endCode);
        excelStruct.setRepeatImportCells(importCells);
    }

    /**
     *
     * 解析一次性导入xml
     *
     * @param excelStruct Excel导入描述文件的结构
     * @param onceList 一次导入集合
     */
    private static void simpleParseOnceImport(ExcelStruct excelStruct, List onceList) {
        if (onceList == null || onceList.size() <= 0) {
            return;
        }
        List<ImportCellDesc> onceImportCells = getImportCellDescList(excelStruct,onceList);
        if (onceImportCells == null || onceImportCells.size() <= 0) {
            return;
        }
        excelStruct.setOnceImportCells(onceImportCells);
    }

    /**
     *
     * 将xml单元格元素解析到ExcelStruct中
     *
     * @param excelStruct Excel导入描述文件的结构
     * @param list 单元格元素list
     */
    private static List<ImportCellDesc> getImportCellDescList(ExcelStruct excelStruct, List list){
        if(list == null || list.size() <= 0){
            return null;
        }
        // 获取CDATA区内的内容
        Element cdataElem = (Element) list.get(0);
        String cdataStr = cdataElem.getTextTrim();
        List<ImportCellDesc> importCells = changeCDATAToList(excelStruct, cdataStr);
        return importCells;
    }

    /**
     * 将CDATA区中的数据转换成我们需要的对象
     */
    private static List<ImportCellDesc> changeCDATAToList(ExcelStruct excelStruct, String cdata) {
        if(StringUtils.isEmpty(cdata)){
            return null;
        }
        // 去掉空白字符
        cdata = cdata.trim().replaceAll("\\s","");
        if(StringUtils.isEmpty(cdata)){
            return null;
        }
        String[] arr = cdata.split(",");
        if (arr == null || arr.length <= 0) {
            return null;
        }
        List<ImportCellDesc> list = new ArrayList<ImportCellDesc>();
        for(int i = 0; i < arr.length; i++){
            if(StringUtils.isEmpty(arr[i])){
                continue;
            }
            String[] kv = arr[i].split("=");
            if(kv.length < 2){
                continue;
            }
            ImportCellDesc cellDesc = new ImportCellDesc();
            cellDesc.setCellRef(kv[0].toUpperCase());
            if(StringUtils.isEmpty(cellDesc.getCellRef())){
                throw new ExcelImportException("xml文件repeatImport有错误");
            }
            cellDesc.setFiledName(kv[1].toUpperCase());
            if(excelStruct != null){
                cellDesc.setValidatorList(excelStruct.getCellValidatorMap().get(cellDesc.getCellRef()));
            }
            list.add(cellDesc);
        }
        return  list;
    }

    /**
     * 读取校验器的相关配置
     */
    private static void parseValidatorConfig(ExcelStruct excelStruct, List validators, List cellValidators) {
        if (excelStruct == null || validators == null || validators.size() <= 0 || cellValidators == null
                || cellValidators.size() <= 0) {
            return;
        }
        // 1.读取校验器的定义
        Element validElem = (Element) validators.get(0);
        if (validElem == null) {
            return;
        }
        List validatorList = validElem.elements("validator");
        if (validatorList == null || validatorList.size() <= 0) {
            return;
        }
        for (Object obj : validatorList) {
            if (obj == null) {
                continue;
            }
            Element validator = (Element) obj;
            String name = validator.attributeValue("name");
            String value = validator.attributeValue("value");
            excelStruct.addSysValidator(name, value);
        }
        // 2.读取单元格的校验器
        Element cellValidElem = (Element) cellValidators.get(0);
        if (cellValidElem == null) {
            return;
        }
        List cellValidatorList = cellValidElem.elements("cell-validator");
        if (cellValidatorList == null || cellValidatorList.size() <= 0) {
            return;
        }
        for (Object obj : cellValidatorList) {
            if (obj == null) {
                continue;
            }
            Element cellValidator = (Element) obj;
            String cellname = cellValidator.attributeValue("cellname"); // 单元格名称
            if (StringUtils.isEmpty(cellname)) {
                continue;
            }
            List cValidators = cellValidator.elements("validator"); // 单元格所使用的校验器
            if (cValidators == null || cValidators.size() <= 0) {
                continue;
            }
            for (Object tmp : cValidators) {
                if (tmp == null) {
                    continue;
                }
                Element validator = (Element) tmp;
                String validatorName = validator.attributeValue("name");
                excelStruct.addCellValidator(cellname, validatorName);
            }
        }
    }


}
