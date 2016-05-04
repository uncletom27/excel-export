package com.meituan.show.settlement.export;

public enum CellStyle {
    NONE("","空"),
    TIME("s=\"1\"","时间");
    
    private final String serializeString;
    private final String desc;

    private CellStyle(String serializeString, String desc) {
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
