/*
 * Copyright (c) 2010-2015 meituan.com
 * All rights reserved.
 * 
 */
package com.meituan.show.settlement.export;

import java.io.IOException;
import java.io.OutputStream;

/**
 * @author xujia06
 * @created 2016年4月28日
 *
 * @version 1.0
 */
public abstract class Excel {
    public static Excel newInstance(OutputStream outputStream) throws IOException{
        return new ExcelImpl(outputStream);
    }
    
    abstract public Excel beginNewSheet(String sheetName) throws IOException;
    abstract public Excel endSheet() throws IOException;
    abstract public Excel addRow(Object obj) throws IOException;
    abstract public Excel addTitle(Class<?> clazz) throws IOException;
    abstract public void finish() throws IOException;
}
