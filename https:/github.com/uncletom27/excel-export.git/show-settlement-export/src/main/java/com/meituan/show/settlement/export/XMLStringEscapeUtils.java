/*
 * Copyright (c) 2010-2015 meituan.com
 * All rights reserved.
 * 
 */
package com.meituan.show.settlement.export;

import java.io.IOException;
import java.io.StringWriter;

/**
 * @author xujia06
 * @created 2016年5月9日
 *
 * @version 1.0
 */
public class XMLStringEscapeUtils {
    private static final char[][] e = {
        "&amp;".toCharArray(),
        "&quot;".toCharArray(), 
        "&lt;".toCharArray(),  
        "&gt;".toCharArray(),
        "&apos;".toCharArray()
    };
    private static final char[] ch = new char[]{
        '&',
        '"', 
        '<',  
        '>',
        '\''
    };
    
    private static char[] getEscape(char c){
        for (int j = 0; j < ch.length; j++) {
            if(ch[j] == c){
                return e[j];
            }
        }
        return null;
    }
    
    public static String escape(String input) throws IOException{
        StringWriter sw = new StringWriter(input.length() * 2);
        char[] inputc = input.toCharArray();
        for (int i = 0; i < inputc.length; i++) {
            char c = inputc[i];
            char[] escape = getEscape(c);
            if(escape != null) {
                sw.write(escape);
            }else {
                sw.write(inputc, i, 1);
            }
        }
        return sw.toString();
    }
}
