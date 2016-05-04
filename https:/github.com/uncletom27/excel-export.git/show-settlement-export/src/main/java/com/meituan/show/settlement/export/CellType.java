package com.meituan.show.settlement.export;

/**
 * https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.cellvalues.aspx
 * @author xujia06
 * @created 2016年4月29日
 *
 * @version 1.0
 */
public enum CellType {
    NONE("",""),
    STRING("t=\"str\"","String. When the item is serialized out as xml, its value is 'str'"),
    BOOLEAN( "t=\"b\"","Boolean. When the item is serialized out as xml, its value is 'b'"),
    DATE("t=\"d\"", "d. When the item is serialized out as xml, its value is 'd'.This item is only available in Office2010"),
    NUMBER("t=\"n\"", "Number. When the item is serialized out as xml, its value is 'n'");
    
    private final String serializeString;
    private final String desc;

    private CellType(String serializeString, String desc) {
        this.serializeString = serializeString;
        this.desc = desc;
    }

    public String getSerializeString() {
        return serializeString;
    }

    public String getDesc() {
        return desc;
    }
    
}
