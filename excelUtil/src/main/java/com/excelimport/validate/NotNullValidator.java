package com.excelimport.validate;

import org.apache.commons.lang3.StringUtils;

/**
 * 不能为空校验器
 */
public class NotNullValidator extends AbstractValidator
{
    public String processValidate()
    {
        if(StringUtils.isEmpty(getFieldValue()))
        {
            return getCellRef() + "单元格数据 : " + getFieldValue() + ", 不可以为空!";
        }
        return OK;
    }
}
