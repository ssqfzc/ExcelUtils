package com.excelimport.bean;

import org.apache.poi.hssf.util.CellReference;

import java.util.ArrayList;
import java.util.List;

/**
 *  单元格的描述信息
 * （1）数据格式校验时，可以精确定位到某个单元格。
 * Created by can on 2017/3/28.
 */
public class ImportCellDesc implements Cloneable{
    /**
     * 引用的单元格；如：A3
     */
    private String cellRef;
    /**
     * 单元格的对应数据库的字段名称；
     * 如：fieldName = "username"
     */
    private String filedName;
    /**
     * 字段值
     */
    private String filedValue;

    /**
     * 字段的校验器
     */
    private List<String> validatorList = new ArrayList<String>();

    public String getCellRef() {
        return cellRef;
    }

    public void setCellRef(String cellRef) {
        this.cellRef = cellRef;
    }

    public String getFiledName() {
        return filedName;
    }

    public void setFiledName(String filedName) {
        this.filedName = filedName;
    }

    public String getFiledValue() {
        return filedValue;
    }

    public void setFiledValue(String filedValue) {
        this.filedValue = filedValue;
    }

    public List<String> getValidatorList() {
        return validatorList;
    }

    public void setValidatorList(List<String> validatorList) {
        this.validatorList = validatorList;
    }

    /**
     * 返回单元格的行标（从1开始）
     * @return	5
     */
    public int getCellRow()
    {
        CellReference ref = new CellReference(cellRef);
        return ref.getRow() + 1;
    }

    /**
     * 返回单元格的列标（从1开始）
     * @return	2
     */
    public int getCellCol()
    {
        CellReference ref = new CellReference(cellRef);
        return ref.getCol() + 1;
    }

    @Override
    public Object clone(){
         ImportCellDesc cellDesc = null;
        try {
            cellDesc = (ImportCellDesc) super.clone();
        } catch (CloneNotSupportedException e) {
            e.printStackTrace();
        }
        if(cellDesc.getValidatorList() != null){
            List<String> a = new ArrayList<String>();
            a.addAll(cellDesc.getValidatorList());
            cellDesc.setValidatorList(a);
        }
        return cellDesc;
    }

    public String toString()
    {
        StringBuffer sb = new StringBuffer(100);
        if(validatorList != null && validatorList.size() > 0)
        {
            sb.append("\tvalidator : [ ");
            for(int i = 0; i < validatorList.size(); i++)
            {
                String validator = validatorList.get(i);
                if(i != validatorList.size() - 1)
                {
                    sb.append(validator).append(" , ");
                }
                else
                {
                    sb.append(validator);
                }
            }
            sb.append(" ]");
        }
        return getCellRef() + "\t(" + getFiledName() + " : " + getFiledValue() + ")" + sb.toString();
    }
}
