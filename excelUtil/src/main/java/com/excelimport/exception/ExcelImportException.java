package com.excelimport.exception;

/**
 * 导入时的异常
 * Created by can on 2017/3/28.
 */
public class ExcelImportException extends RuntimeException{
    private static final long serialVersionUID = 8036129856262757376L;

    private String msg;

    public ExcelImportException() {}

    public ExcelImportException(String message) {
        super(message);
        this.msg = message;
    }

    public ExcelImportException(String message, Throwable cause) {
        super(message, cause);
        this.msg = message;
    }

    public String getMsg()
    {
        return msg;
    }

    public void setMsg(String msg)
    {
        this.msg = msg;
    }
}
