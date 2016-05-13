/*
 * Copyright (c) 2010-2015 meituan.com
 * All rights reserved.
 * 
 */
package com.meituan.show.settlement.export;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.net.ServerSocket;
import java.net.Socket;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;
import java.util.prefs.Preferences;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * @author xujia06
 * @created 2016年4月28日
 *
 * @version 1.0
 */
class ExcelImpl extends Excel{
    private static final String Content_Types_path = "[Content_Types].xml";
    private static final String Content_Types_head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>";
    private static final String Content_Types_tail = "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>";       
    private static final String Content_types_body_template = "<Override PartName=\"/xl/worksheets/sheet%s.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";        
    
    private static final String _rels_rels_path = "_rels/.rels";
    private static final String _rels_rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";
    
    private static final String docProps_app_path = "docProps/app.xml";
    private static final String docProps_app_head_template = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>%s</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=\"%s\" baseType=\"lpstr\">";
    private static final String docProps_app_tail = "</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0300</AppVersion></Properties>";
    private static final String docProps_app_body_template = "<vt:lpstr>%s</vt:lpstr>";
    
    private static final String docProps_core_path = "docProps/core.xml";
    private static final String docProps_core_body_template = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator></dc:creator><dcterms:created xsi:type=\"dcterms:W3CDTF\">2006-09-16T00:00:00Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2016-04-27T11:36:24Z</dcterms:modified></cp:coreProperties>";
    
    private static final String xl_workbook_path = "xl/workbook.xml";
    private static final String xl_workbook_head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion appName=\"xl\" lastEdited=\"5\" lowestEdited=\"4\" rupBuild=\"9302\"/><workbookPr filterPrivacy=\"1\" defaultThemeVersion=\"124226\"/><bookViews><workbookView xWindow=\"240\" yWindow=\"105\" windowWidth=\"14805\" windowHeight=\"8010\"/></bookViews><sheets>";
    private static final String xl_workbook_tail = "</sheets><calcPr calcId=\"122211\"/></workbook>";
    private static final String xl_workbook_body_template= "<sheet name=\"%s\" sheetId=\"%s\" r:id=\"rId%s\"/>";
    
    private static final String xl_styles_path = "xl/styles.xml";
    private static final String xl_styles = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><fonts count=\"2\" x14ac:knownFonts=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"宋体\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"9\"/><name val=\"宋体\"/><family val=\"3\"/><charset val=\"134\"/><scheme val=\"minor\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"常规\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleMedium9\"/><extLst><ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/></ext></extLst></styleSheet>";
    
    private static final String xl_theme_theme1_path = "xl/theme/theme1.xml";
    private static final String xl_theme_theme1 = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office 主题​​\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2><a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Cambria\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>";
    
    private static final String xl_rels_workbook_path = "xl/_rels/workbook.xml.rels";
    private static final String xl_rels_workbook_head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    private static final String xl_rels_workbook_tail_template = "<Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/></Relationships>";
    private static final String xl_rels_workbook_body_template = "<Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%s.xml\"/>";
    
    private static final String xl_printerSettings_printerSettings_path_template = "xl/printerSettings/printerSettings%s.bin";
    private static final String xl_printerSettings_printerSerttings_body_base64 = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAAbcAAAAU/+AAwEACQCaCzQIZAABAA8AWAICAAEAWAIDAAEAQQA0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";
    private static final byte[] xl_printerSettings_printerSettings_Body = Base64.base64ToByteArray(xl_printerSettings_printerSerttings_body_base64);
    
    private static final String xl_worksheets_sheet_path_template = "xl/worksheets/sheet%s.xml";
    private static final String xl_worksheets_sheet_head_template = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><dimension ref=\"%s\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"13.5\" x14ac:dyDescent=\"0.15\"/><sheetData>";
    private static final String xl_worksheets_sheet_tail = "</sheetData><phoneticPr fontId=\"1\" type=\"noConversion\" /><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\" /><pageSetup paperSize=\"9\" orientation=\"portrait\" r:id=\"rId1\" /></worksheet>";
    private static final String xl_worksheets_sheet_body_row_begin_template = "<row r=\"%s\" spans=\"1:1\" x14ac:dyDescent=\"0.15\">";
    private static final String xl_worksheets_sheet_body_row_end = "</row>";
    private static final String xl_worksheets_sheet_body_sell_template = "<c r=\"%s\" %s %s><v>%s</v></c>";
    
    private static final String xl_worksheets_rels_sheet_path_template = "xl/worksheets/_rels/sheet%s.xml.rels";
    private static final String xl_worksheets_rels_sheet_template = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" Target=\"../printerSettings/printerSettings%s.bin\"/></Relationships>";
    
    
    private static final String[] ch = new String[]{"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
    private ZipOutputStream zipOutputStream;
    private List<String> sheetNames = new ArrayList<>();
    private long currentRow = 1l;
    private int currentClumn = 1;
    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    
    public ExcelImpl(OutputStream outputStream) throws IOException {
        this.zipOutputStream = new ZipOutputStream(outputStream);
        write_rels_rels();
        write_docProps_core();
        write_xl_styles();
        write_xl_theme_theme1();
    }

    @Override
    public Excel beginNewSheet(String sheetName) throws IOException {
        sheetNames.add(sheetName);
        
        ZipEntry xl_worksheets_sheet = new ZipEntry(String.format(xl_worksheets_sheet_path_template, this.sheetNames.size()));
        zipOutputStream.putNextEntry(xl_worksheets_sheet);
        zipOutputStream.write(String.format(xl_worksheets_sheet_head_template, "A1").getBytes("UTF-8"));
        
        return this;
    }

    @Override
    public Excel endSheet() throws IOException {
        zipOutputStream.write(xl_worksheets_sheet_tail.getBytes("UTF-8"));
        zipOutputStream.closeEntry();
        this.currentRow = 1L;
        this.currentClumn = 1;
        
        return this;
    }

    @Override
    public Excel addRow(Object obj) throws UnsupportedEncodingException, IOException {
        this.beginNewRow();
        
        List<FieldMeta> fieldMetas = getMeta(obj.getClass()).fieldMetas;
        for (FieldMeta fieldMeta : fieldMetas) {
            Object value = fieldMeta.valueOf(obj);
            if(value == null){
                this.addCell( CellType.STRING.getSerializeString(), CellStyle.NONE.getSerializeString(), null);
                continue;
            }
            if(value instanceof String){
                this.addCell( CellType.STRING.getSerializeString(), CellStyle.NONE.getSerializeString(), XMLStringEscapeUtils.escape((String)value));
            }else if(value instanceof Date){
                this.addCell( CellType.STRING.getSerializeString(), CellStyle.NONE.getSerializeString(), (simpleDateFormat.format((Date) value)));
            }else if(value instanceof Boolean){
                this.addCell( CellType.BOOLEAN.getSerializeString(), CellStyle.NONE.getSerializeString(), (Boolean)value.equals(true)? 1 : 0);
            }else if(value instanceof Number){
                this.addCell( CellType.NUMBER.getSerializeString(), CellStyle.NONE.getSerializeString(), value);
            }else {
                throw new RuntimeException(String.format("不支持的数据类型%s", value.getClass()));
            }
        }
        
        this.endRow();
        return this;
    }
    
    @Override
    public Excel addTitle(Class<?> clazz) throws IOException {
        this.beginNewRow();
        List<FieldMeta> fieldMetas = getMeta(clazz).fieldMetas;
        for (FieldMeta fieldMeta : fieldMetas) {
            this.addCell( CellType.STRING.getSerializeString(), CellStyle.NONE.getSerializeString(), fieldMeta.title);
        }
        this.endRow();
        return this;
    }
    
    @Override
    public void finish() throws UnsupportedEncodingException, IOException {
        write_Content_Types();
        write_docProps_app();
        write_xl_workbook();
        write_xl_rels_workbook();
        write_xl_printerSettings_printerSettings();
        write_xl_worksheets_rels_sheet();
        this.zipOutputStream.finish();
    }
    
    private Excel beginNewRow() throws UnsupportedEncodingException, IOException{
        String row = String.format(xl_worksheets_sheet_body_row_begin_template, this.currentRow);
        this.zipOutputStream.write(row.getBytes("UTF-8"));
        return this;
    }
    
    private Excel endRow() throws UnsupportedEncodingException, IOException{
        this.zipOutputStream.write(xl_worksheets_sheet_body_row_end.getBytes("UTF-8"));
        this.currentRow = this.currentRow+1;
        this.currentClumn = 1;
        return this;
    }
    
    private Excel addCell(String type, String style, Object vlaue) throws UnsupportedEncodingException, IOException{
        if(vlaue != null){
            String cell = String.format(xl_worksheets_sheet_body_sell_template, this.getCellName() , type, style, vlaue);
            this.zipOutputStream.write(cell.getBytes("UTF-8"));
        }
        this.currentClumn = this.currentClumn + 1;
        return this;
    }

    private void write_xl_worksheets_rels_sheet() throws IOException, UnsupportedEncodingException {
        for (int i = 0; i < this.sheetNames.size(); i++) {
            ZipEntry xl_worksheets_rels_sheet = new ZipEntry(String.format(xl_worksheets_rels_sheet_path_template, i+1));
            zipOutputStream.putNextEntry(xl_worksheets_rels_sheet);
            zipOutputStream.write(String.format(ExcelImpl.xl_worksheets_rels_sheet_template, i+1).getBytes("UTF-8"));//TODO
            zipOutputStream.closeEntry();
        }
    }

    private void write_xl_printerSettings_printerSettings() throws IOException {
        for (int i = 0; i < this.sheetNames.size(); i++) {
            ZipEntry xl_printerSettings_printerSettings = new ZipEntry(String.format(xl_printerSettings_printerSettings_path_template, i+1));
            zipOutputStream.putNextEntry(xl_printerSettings_printerSettings);
            zipOutputStream.write(xl_printerSettings_printerSettings_Body);
            zipOutputStream.closeEntry();
        }
    }

    private void write_xl_rels_workbook() throws IOException, UnsupportedEncodingException {
        ZipEntry xl_rels_workbook = new ZipEntry(xl_rels_workbook_path);
        zipOutputStream.putNextEntry(xl_rels_workbook);
        zipOutputStream.write(ExcelImpl.xl_rels_workbook_head.getBytes("UTF-8"));
        for (int i = 0; i < this.sheetNames.size(); i++) {
            zipOutputStream.write(String.format(xl_rels_workbook_body_template, i+1, i+1).getBytes("UTF-8"));
        }
        zipOutputStream.write(String.format(ExcelImpl.xl_rels_workbook_tail_template, this.sheetNames.size()+1, sheetNames.size()+2).getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }

    private void write_xl_workbook() throws IOException, UnsupportedEncodingException {
        ZipEntry xl_workbook = new ZipEntry(xl_workbook_path);
        zipOutputStream.putNextEntry(xl_workbook);
        zipOutputStream.write(ExcelImpl.xl_workbook_head.getBytes("UTF-8"));
        for (int i = 0; i < this.sheetNames.size(); i++) {
            zipOutputStream.write(String.format(xl_workbook_body_template, this.sheetNames.get(i), i+1, i+1).getBytes("UTF-8"));
        }
        zipOutputStream.write(ExcelImpl.xl_workbook_tail.getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }

    private void write_docProps_app() throws IOException, UnsupportedEncodingException {
        ZipEntry docProps_app = new ZipEntry(docProps_app_path);
        zipOutputStream.putNextEntry(docProps_app);
        zipOutputStream.write(String.format(ExcelImpl.docProps_app_head_template, this.sheetNames.size(), this.sheetNames.size()).getBytes("UTF-8"));
        for (String sheetName : this.sheetNames) {
            zipOutputStream.write(String.format(docProps_app_body_template, sheetName).getBytes("UTF-8"));
        }
        zipOutputStream.write(ExcelImpl.docProps_app_tail.getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }

    private void write_Content_Types() throws IOException, UnsupportedEncodingException {
        ZipEntry Content_Types = new ZipEntry(Content_Types_path);
        zipOutputStream.putNextEntry(Content_Types);
        zipOutputStream.write(ExcelImpl.Content_Types_head.getBytes("UTF-8"));
        for (int i = 0; i < this.sheetNames.size(); i++) {
            zipOutputStream.write(String.format(Content_types_body_template, i+1).getBytes("UTF-8"));
        }
        zipOutputStream.write(ExcelImpl.Content_Types_tail.getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }
    
    private void write_xl_theme_theme1() throws IOException, UnsupportedEncodingException {
        ZipEntry xl_theme_theme1 = new ZipEntry(xl_theme_theme1_path);
        zipOutputStream.putNextEntry(xl_theme_theme1);
        zipOutputStream.write(ExcelImpl.xl_theme_theme1.getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }

    private void write_xl_styles() throws IOException, UnsupportedEncodingException {
        ZipEntry xl_styles = new ZipEntry(xl_styles_path);
        zipOutputStream.putNextEntry(xl_styles);
        zipOutputStream.write(ExcelImpl.xl_styles.getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }

    private void write_docProps_core() throws IOException, UnsupportedEncodingException {
        ZipEntry docProps_core = new ZipEntry(docProps_core_path);
        zipOutputStream.putNextEntry(docProps_core);
        zipOutputStream.write(String.format(ExcelImpl.docProps_core_body_template,"", "").getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }

    private void write_rels_rels() throws IOException, UnsupportedEncodingException {
        ZipEntry _rels_rels = new ZipEntry(_rels_rels_path);
        zipOutputStream.putNextEntry(_rels_rels);
        zipOutputStream.write(ExcelImpl._rels_rels.getBytes("UTF-8"));
        zipOutputStream.closeEntry();
    }
    
    private String getCellName(){
        String s = "";
        int n = this.currentClumn;
        do {
            int j = n%26;
            n = n/26;
            if(j == 0){
                j = 26;
                n = n-1;
            }
            s = ch[j-1] + s;
        } while (n != 0);
        
        return s + this.currentRow;
    }
    
    protected static final ConcurrentHashMap<Class<?>, ClassMeta> metasMap = new ConcurrentHashMap<>();
    
    public static ClassMeta getMeta(Class<?> clazz) {
        ClassMeta result = metasMap.get(clazz);
        if (result != null) {
            return result;
        }
        result = new ClassMeta();

        List<FieldMeta> fieldMetas = new LinkedList<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            Cell cell = field.getAnnotation(Cell.class);
            FieldMeta m = new FieldMeta();
            m.field = field;
            m.field.setAccessible(true);
            if (cell == null) {
                throw new RuntimeException(String.format("类%s的%s字段没有Cell注解", clazz.getName(), field.getName()));
            } else {
                m.type = cell.type();
                m.style = cell.style();
                String title = cell.title();
                if(title.equals("")){
                    m.title = field.getName();
                }else {
                    m.title = title;
                }
                m.column = cell.order();
            }
            fieldMetas.add(m);
        }
        Collections.sort(fieldMetas, new Comparator<FieldMeta>(){

            @Override
            public int compare(FieldMeta o1, FieldMeta o2) {
                int i = o1.column - o2.column;
                if(i != 0){
                    return i;
                }else {
                    return o1.field.getName().compareTo(o1.field.getName());
                }
            }
            
        });
        result.fieldMetas = fieldMetas;
        metasMap.put(clazz, result);
        return result;
    }

    @SuppressWarnings("unused")
    private static class FieldMeta {

        private Field field;
        private CellType type;
        private CellStyle style;
        private String title;
        private int column;

        private Object valueOf(Object o) {
            try {
                return field.get(o);
            } catch (IllegalArgumentException | IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private static class ClassMeta {
        private List<FieldMeta> fieldMetas;
    }
    
    public static void main(String[] args) throws IOException {
        Server server = new Server();
        server.start();
        Socket s = new Socket("localhost", 8090);
        OutputStream fos = s.getOutputStream();
        Excel exel = new ExcelImpl(fos);
        exel.beginNewSheet("付款表");
        exel.addTitle(Test.class);
        for (int i = 0; i < 1000000000; i++) {
            Test t = new Test();
            exel.addRow(t);
        }
        exel.endSheet();
        exel.finish();
        s.close();
        System.err.println("finished");
    }
    
    private static class Server extends Thread {

        @Override
        public void run() {
            long l = 0;
            ServerSocket ss;
            try (FileOutputStream fos = new FileOutputStream("aaa.xlsx")){
                ss = new ServerSocket(8090);
                Socket s = ss.accept();
                InputStream inputStream = s.getInputStream();
                byte[] buff = new byte[1024];
                int read = -1;
                while((read = inputStream.read(buff)) != -1){
                    l = l+ read;
                    System.err.println(l);
                    fos.write(buff, 0, read);
                }
                
            } catch (IOException e) {
                e.printStackTrace();
            }
            
        }
        
    }
    
    private static class Test {
        @Cell(order = 1)
        Date d = new Date();
        @Cell(order = 2)
        private String s = "大法师打发是发送到发送到发送到发送到发你好阿萨斯的发送";
        @Cell(order = 2)
        private long l = 123l;
        @Cell(order = 3)
        private boolean b = true;
        @Cell(order = 4)
        private  float f = 123.2f;
        @Cell(order = 4)
        private double dd = 123.2d;
        
    }
    
    
    
    
    /**
     * copy from jdk
     * Static methods for translating Base64 encoded strings to byte arrays
     * and vice-versa.
     *
     * @author  Josh Bloch
     * @see     Preferences
     * @since   1.4
     */
    @SuppressWarnings("unused")
    private static class Base64 {
        /**
         * Translates the specified byte array into a Base64 string as per
         * Preferences.put(byte[]).
         */
        static String byteArrayToBase64(byte[] a) {
            return byteArrayToBase64(a, false);
        }

        /**
         * Translates the specified byte array into an "alternate representation"
         * Base64 string.  This non-standard variant uses an alphabet that does
         * not contain the uppercase alphabetic characters, which makes it
         * suitable for use in situations where case-folding occurs.
         */
        static String byteArrayToAltBase64(byte[] a) {
            return byteArrayToBase64(a, true);
        }

        private static String byteArrayToBase64(byte[] a, boolean alternate) {
            int aLen = a.length;
            int numFullGroups = aLen/3;
            int numBytesInPartialGroup = aLen - 3*numFullGroups;
            int resultLen = 4*((aLen + 2)/3);
            StringBuffer result = new StringBuffer(resultLen);
            char[] intToAlpha = (alternate ? intToAltBase64 : intToBase64);

            // Translate all full groups from byte array elements to Base64
            int inCursor = 0;
            for (int i=0; i<numFullGroups; i++) {
                int byte0 = a[inCursor++] & 0xff;
                int byte1 = a[inCursor++] & 0xff;
                int byte2 = a[inCursor++] & 0xff;
                result.append(intToAlpha[byte0 >> 2]);
                result.append(intToAlpha[(byte0 << 4)&0x3f | (byte1 >> 4)]);
                result.append(intToAlpha[(byte1 << 2)&0x3f | (byte2 >> 6)]);
                result.append(intToAlpha[byte2 & 0x3f]);
            }

            // Translate partial group if present
            if (numBytesInPartialGroup != 0) {
                int byte0 = a[inCursor++] & 0xff;
                result.append(intToAlpha[byte0 >> 2]);
                if (numBytesInPartialGroup == 1) {
                    result.append(intToAlpha[(byte0 << 4) & 0x3f]);
                    result.append("==");
                } else {
                    // assert numBytesInPartialGroup == 2;
                    int byte1 = a[inCursor++] & 0xff;
                    result.append(intToAlpha[(byte0 << 4)&0x3f | (byte1 >> 4)]);
                    result.append(intToAlpha[(byte1 << 2)&0x3f]);
                    result.append('=');
                }
            }
            // assert inCursor == a.length;
            // assert result.length() == resultLen;
            return result.toString();
        }

        /**
         * This array is a lookup table that translates 6-bit positive integer
         * index values into their "Base64 Alphabet" equivalents as specified
         * in Table 1 of RFC 2045.
         */
        private static final char intToBase64[] = {
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
            'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
            'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
            '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '+', '/'
        };

        /**
         * This array is a lookup table that translates 6-bit positive integer
         * index values into their "Alternate Base64 Alphabet" equivalents.
         * This is NOT the real Base64 Alphabet as per in Table 1 of RFC 2045.
         * This alternate alphabet does not use the capital letters.  It is
         * designed for use in environments where "case folding" occurs.
         */
        private static final char intToAltBase64[] = {
            '!', '"', '#', '$', '%', '&', '\'', '(', ')', ',', '-', '.', ':',
            ';', '<', '>', '@', '[', ']', '^',  '`', '_', '{', '|', '}', '~',
            'a', 'b', 'c', 'd', 'e', 'f', 'g',  'h', 'i', 'j', 'k', 'l', 'm',
            'n', 'o', 'p', 'q', 'r', 's', 't',  'u', 'v', 'w', 'x', 'y', 'z',
            '0', '1', '2', '3', '4', '5', '6',  '7', '8', '9', '+', '?'
        };

        /**
         * Translates the specified Base64 string (as per Preferences.get(byte[]))
         * into a byte array.
         *
         * @throw IllegalArgumentException if <tt>s</tt> is not a valid Base64
         *        string.
         */
        static byte[] base64ToByteArray(String s) {
            return base64ToByteArray(s, false);
        }

        /**
         * Translates the specified "alternate representation" Base64 string
         * into a byte array.
         *
         * @throw IllegalArgumentException or ArrayOutOfBoundsException
         *        if <tt>s</tt> is not a valid alternate representation
         *        Base64 string.
         */
        static byte[] altBase64ToByteArray(String s) {
            return base64ToByteArray(s, true);
        }

        private static byte[] base64ToByteArray(String s, boolean alternate) {
            byte[] alphaToInt = (alternate ?  altBase64ToInt : base64ToInt);
            int sLen = s.length();
            int numGroups = sLen/4;
            if (4*numGroups != sLen)
                throw new IllegalArgumentException(
                    "String length must be a multiple of four.");
            int missingBytesInLastGroup = 0;
            int numFullGroups = numGroups;
            if (sLen != 0) {
                if (s.charAt(sLen-1) == '=') {
                    missingBytesInLastGroup++;
                    numFullGroups--;
                }
                if (s.charAt(sLen-2) == '=')
                    missingBytesInLastGroup++;
            }
            byte[] result = new byte[3*numGroups - missingBytesInLastGroup];

            // Translate all full groups from base64 to byte array elements
            int inCursor = 0, outCursor = 0;
            for (int i=0; i<numFullGroups; i++) {
                int ch0 = base64toInt(s.charAt(inCursor++), alphaToInt);
                int ch1 = base64toInt(s.charAt(inCursor++), alphaToInt);
                int ch2 = base64toInt(s.charAt(inCursor++), alphaToInt);
                int ch3 = base64toInt(s.charAt(inCursor++), alphaToInt);
                result[outCursor++] = (byte) ((ch0 << 2) | (ch1 >> 4));
                result[outCursor++] = (byte) ((ch1 << 4) | (ch2 >> 2));
                result[outCursor++] = (byte) ((ch2 << 6) | ch3);
            }

            // Translate partial group, if present
            if (missingBytesInLastGroup != 0) {
                int ch0 = base64toInt(s.charAt(inCursor++), alphaToInt);
                int ch1 = base64toInt(s.charAt(inCursor++), alphaToInt);
                result[outCursor++] = (byte) ((ch0 << 2) | (ch1 >> 4));

                if (missingBytesInLastGroup == 1) {
                    int ch2 = base64toInt(s.charAt(inCursor++), alphaToInt);
                    result[outCursor++] = (byte) ((ch1 << 4) | (ch2 >> 2));
                }
            }
            // assert inCursor == s.length()-missingBytesInLastGroup;
            // assert outCursor == result.length;
            return result;
        }

        /**
         * Translates the specified character, which is assumed to be in the
         * "Base 64 Alphabet" into its equivalent 6-bit positive integer.
         *
         * @throw IllegalArgumentException or ArrayOutOfBoundsException if
         *        c is not in the Base64 Alphabet.
         */
        private static int base64toInt(char c, byte[] alphaToInt) {
            int result = alphaToInt[c];
            if (result < 0)
                throw new IllegalArgumentException("Illegal character " + c);
            return result;
        }

        /**
         * This array is a lookup table that translates unicode characters
         * drawn from the "Base64 Alphabet" (as specified in Table 1 of RFC 2045)
         * into their 6-bit positive integer equivalents.  Characters that
         * are not in the Base64 alphabet but fall within the bounds of the
         * array are translated to -1.
         */
        private static final byte base64ToInt[] = {
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, 62, -1, -1, -1, 63, 52, 53, 54,
            55, 56, 57, 58, 59, 60, 61, -1, -1, -1, -1, -1, -1, -1, 0, 1, 2, 3, 4,
            5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
            24, 25, -1, -1, -1, -1, -1, -1, 26, 27, 28, 29, 30, 31, 32, 33, 34,
            35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51
        };

        /**
         * This array is the analogue of base64ToInt, but for the nonstandard
         * variant that avoids the use of uppercase alphabetic characters.
         */
        private static final byte altBase64ToInt[] = {
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 0, 1,
            2, 3, 4, 5, 6, 7, 8, -1, 62, 9, 10, 11, -1 , 52, 53, 54, 55, 56, 57,
            58, 59, 60, 61, 12, 13, 14, -1, 15, 63, 16, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, 17, -1, 18, 19, 21, 20, 26, 27, 28, 29, 30, 31, 32, 33,
            34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50,
            51, 22, 23, 24, 25
        };

        public static void main(String args[]) {
            String s = Base64.byteArrayToBase64(xl_printerSettings_printerSettings_Body);
            System.err.println(s);
            byte[] bs = Base64.base64ToByteArray(s);
            for (int i = 0; i < bs.length; i++) {
                if(xl_printerSettings_printerSettings_Body[i]!=bs[i]){
                    System.err.println("wrong");
                }
            }
            System.err.println("finished");
        }
    }

}
