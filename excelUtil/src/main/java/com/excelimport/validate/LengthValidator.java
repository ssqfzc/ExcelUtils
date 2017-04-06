package com.excelimport.validate;

import org.apache.commons.lang3.StringUtils;

/**
 * 长度校验
 */
public class LengthValidator extends AbstractValidator {

    @Override
    public String processValidate() {
        int maxLength = 100;
        int minLength = 6;

        if(StringUtils.isNotEmpty(getFieldValue()) && getFieldValue().length() >= minLength && getFieldValue().length() <= maxLength)
        {
            return OK;
        }

        return getCellRef() + "单元格数据 : " + getFieldValue() + ", 长度不合法, 必须在 " + minLength + "~" + maxLength + " 之间!";
    }
}
