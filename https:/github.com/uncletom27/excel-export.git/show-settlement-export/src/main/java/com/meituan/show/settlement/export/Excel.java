/*
 * Copyright (c) 2010-2015 meituan.com
 * All rights reserved.
 * 
 */
package com.meituan.show.settlement.export;

import java.io.IOException;

/**
 * @author xujia06
 * @created 2016年4月28日
 *
 * @version 1.0
 */
public interface Excel {
    Excel beginNewSheet(String sheetName) throws IOException;
    Excel endSheet() throws IOException;
    Excel addRow(Object obj) throws IOException;
    Excel addTitle(Class<?> clazz) throws IOException;
    void finish() throws IOException;
}
