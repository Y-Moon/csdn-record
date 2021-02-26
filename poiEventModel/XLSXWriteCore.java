package com.zhibeitech.common.utils;
import com.zhibeitech.modules.encryption.controller.Model;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FontScheme;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Set;
import java.util.regex.Pattern;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * @author Y-Moon
 * <p>Description: xlsx文件的处理</p>
 */
public class XLSXWriteCore {
    /**
     * <p>Field in: 进来的表路径</p>
     * <p>Field out: 返回的表路径</p>
     * <p>Field workbook: 工作簿</p>
     * <p>Field decNum: xlsx文件处理个数记录</p>
     */
    private final String in;
    private final String out;
    private SXSSFWorkbook workbook;
    private int decNum;
    /**
     * <p>Field stylesTable: 单元格样式</p>
     */
    public static StylesTable stylesTable;
    public XLSXWriteCore(String in, String out) throws FileNotFoundException {
        this.in = in;
        this.out = out;
        //根据硬件配置，调整一次读写的行数
        this.workbook = new SXSSFWorkbook(5000);
    }
    /**
     * <p>Description: 处理单个sheet(本案例调用此方法)</p>
     *
     * @throws Exception 异常
     */
    public void processOneSheet() throws Exception {
        OPCPackage pkg = OPCPackage.open(in);
        XSSFReader r = new XSSFReader(pkg);
        //获得excel的共享字符串表
        SharedStringsTable sst = r.getSharedStringsTable();
        //获得excel的样式表
        stylesTable = r.getStylesTable();
        //解析器的创建
        XMLReader parser = fetchSheetParser(sst);
        //只有一张表的话默认id为rId1
        InputStream sheet = r.getSheet("rId1");
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        //表写入
        FileOutputStream outputStream = new FileOutputStream(out);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        workbook.dispose();
        workbook.close();
        sheet.close();
        System.out.println("总计：" + decNum);
    }

    /**
     * <p>Description: 处理所有sheet</p>
     *
     * @throws Exception 异常
     */
    public void  processAllSheets() throws Exception {
        OPCPackage pkg = OPCPackage.open(in);
        XSSFReader r = new XSSFReader(pkg);
        //获得excel的共享字符串表
        SharedStringsTable sst = r.getSharedStringsTable();
        //获得excel的样式表
        stylesTable = r.getStylesTable();
        //解析器的创建
        XMLReader parser = fetchSheetParser(sst);
        //多表解析
        Iterator<InputStream> sheets = r.getSheetsData();
        while (sheets.hasNext()) {
            System.out.println("Processing new sheet:\n");
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
        //表写入
        FileOutputStream outputStream = new FileOutputStream(out);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        workbook.dispose();
        workbook.close();
    }

    /**
     * <p>Description: 获取XML访问对象</p>
     *
     * @param sst 共享字符串表对象
     * @return XMLReader XML访问对象
     * @throws SAXException SAX异常
     */
    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("com.sun.org.apache.xerces.internal.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * <p>ClassName: SheetHandler</p>
     * <p>Description: sheet处理类</p>
     * <p>Author: sloth</p>
     * <p>Date: 2020-12-06</p>
     */
    private class SheetHandler extends DefaultHandler {
        /**
         * <p>Field sheet: 表对象</p>
         * <p>Field SXSSFRow: 行对象</p>
         * <p>Field cell: 单元格对象</p>
         */
        private SXSSFSheet sheet;
        private SXSSFRow row;
        private SXSSFCell cell;
        /**
         * <p>Field styleMap: 样式存储库对象</p>
         * <p>Field fontMap: 字体存储库对象</p>
         */
        //样式映射仓库，如果样式太多，可能导致这个hashmap变成超大对象从而无法释放，引起oom(一般不会)
        private HashMap<Integer, XSSFCellStyle> styleMap;
        private HashMap<Integer, XSSFFont> fontMap;
        /**
         * <p>Field rowNumber: 行坐标</p>
         * <p>Field columnNumber: 列坐标</p>
         */
        private int rowNumber;
        private int columnNumber;
        /**
         * <p>Field defaultColumnWidth: 缺省行宽</p>
         * <p>Field defaultRowHeight: 缺省列宽</p>
         */
        private double defaultColumnWidth;
        private double defaultRowHeight;
        /**
         * <p>Field logger: 日志</p>
         */
//        private static Logger logger = Logger.getLogger(SheetHandler.class);
        /**
         * <p>Field sst: 共享字符串表对象</p>
         */
        private SharedStringsTable sst;
        /**
         * <p>Field lastContents: 单元格内容</p>
         */
        private String lastContents;
        /**
         * <p>Field nextIsString: 是否是字符串</p>
         */
        private boolean nextIsString;

        /**
         * <p>Description: 构造函数初始化</p>
         *
         * @param sst 共享字符串表对象
         */
        private SheetHandler(SharedStringsTable sst) {
            //init
            this.sst = sst;
            styleMap = new HashMap<>();
            fontMap = new HashMap<>();
        }

        /**
         * <p>Title: startElement</p>
         * <p>Description: </p>
         *
         * @param uri        uri
         * @param localName  localName
         * @param name       XML标签名
         * @param attributes XML标签对象
         * @throws SAXException SAX异常
         * @see org.xml.sax.helpers.DefaultHandler#startElement(java.lang.String, java.lang.String, java.lang.String, org.xml.sax.Attributes)
         */
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            //init
            switch (name) {
                case "c": {
                    stylesTable.getNumCellStyles();
                    //xml文件中单元格开始标签
                    cell = row.createCell(columnNumber);
                    String cellFormat = attributes.getValue("t");
                    String cellStyleStr = attributes.getValue("s");
                    nextIsString = "s".equals(cellFormat);
                    //为单元格设置样式
                    if (null != cellStyleStr) {
                        //取出单元格样式
                        int styleIndex = Integer.parseInt(cellStyleStr);
                        XSSFCellStyle oldCellStyle = stylesTable.getStyleAt(styleIndex);
                        //根据旧表为新表创建单元格样式，将样式存入样式hashMap中
                        XSSFCellStyle newCellStyle = styleMap.get(styleIndex);
                        if (null == newCellStyle) {
                            newCellStyle =(XSSFCellStyle) workbook.createCellStyle();
                            newCellStyle.cloneStyleFrom(oldCellStyle);
//                            newCellStyle.getFont().setScheme(FontScheme.NONE);
//                            newCellStyle.getFont().setFamily(2);
                            styleMap.put(styleIndex, newCellStyle);
                        }
                        cell.setCellStyle(newCellStyle);
                    }
                    break;
                }
                case "row": {
                    //Xml文件中行标签开始时，初始化columnNumber
                    columnNumber = 0;
                    String ht = attributes.getValue("ht");
                    row = sheet.createRow(rowNumber - 1);
                    if (null != ht) {
                        double h = Double.parseDouble(ht);
                        row.setHeight((short) (h * 20));
                    }
                    break;
                }
                case "sheetFormatPr": {
                    //xml文件中的表格属性标签
                    String RowHeightStr = attributes.getValue("defaultRowHeight");
                    String ColWidthStr = attributes.getValue("defaultColWidth");
                    if (null != RowHeightStr) {
                        defaultRowHeight = Double.parseDouble(RowHeightStr);
                        System.out.println("defaultRowHeight = " + defaultRowHeight);
                    }
                    if (null != ColWidthStr) {
                        defaultColumnWidth = Double.parseDouble(ColWidthStr);
                        System.out.println("defaultColumnWidth = " + defaultColumnWidth);
                    }
                    break;
                }
                case "worksheet": {
                    //xml文件中表格数据开始标签
                    sheet = workbook.createSheet();
                    rowNumber = 1;
                    columnNumber = 0;
                    defaultRowHeight = 0;
                    defaultColumnWidth = 0;
                    break;
                }
                case "col": {
                    //xml文件中的列属性标签
                    String widths = attributes.getValue("width");
                    String maxs = attributes.getValue("max");
                    String mins = attributes.getValue("min");
                    int width = (int) Double.parseDouble(widths);
                    int max = (int) Double.parseDouble(maxs);
                    int min = (int) Double.parseDouble(mins);
                    for (int i = min; i <= max; i++) {
                        sheet.setColumnWidth(i - 1, width * 256);
                    }
                    break;
                }
                case "mergeCell": {
                    //合并单元格标签
                    String[] refs = attributes.getValue("ref").split(":");
                    int[] pre = new int[2];
                    int[] behind = new int[2];
                    for (int i = 0; i < 2; i++) {
                        //将单元格的列数字符解析为数字
                        String[] behindString = new String[2];
                        behindString[i] = refs[i].replaceAll("[\\d]*", "");
                        for (int k = 0; k < behindString[i].length(); k++) {
                            pre[i] = pre[i] * 26 + (behindString[i].charAt(k) - 'A') + 1;
                        }
                        pre[i] -= 1;
                        //讲单元格的行数字符解析为数字
                        behind[i] = Integer.parseInt(refs[i].replaceAll("[\\D]*", "")) - 1;
                    }
                    sheet.addMergedRegion(new CellRangeAddress(behind[0], behind[1], pre[0], pre[1]));
                    break;
                }
            }
            lastContents = "";
        }

        /**
         * <p>Title: endElement</p>
         * <p>Description: </p>
         *
         * @param uri       uri
         * @param localName localName
         * @param name      XML标签名
         * @throws SAXException SAX异常
         * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String, java.lang.String, java.lang.String)
         */

        public void endElement(String uri, String localName, String name) throws SAXException {
            // v => contents of a cell
            // Output after we've seen the string contents
            switch (name) {
                case "c": {
                    ++columnNumber;
                    break;
                }
                //存在value时才进行设置
                case "v": {
                    if (nextIsString) {
                        //将共享字符串序列转化
                        int idx = Integer.parseInt(lastContents);
                        lastContents = sst.getItemAt(idx).getString();
                        cell.setCellType(CellType.STRING);
                        nextIsString = false;
                    }

                    //普通Str的处理
                    dealStr(lastContents);
                    cell.setCellValue(lastContents);
//                    cell.setCellValue(Double.parseDouble(lastContents));
                    break;
                }
                case "t":{
                    //inline的Str处理
                    dealStr(lastContents);
                    break;
                }
                case "row": {
                    System.out.println();
                    //标题只输出一次
                    insert();
                    break;
                }
            }
        }

        /**
         * <p>Description: 处理方法</p>
         *
         * @param encryptedStr 对单元格需要处理的字符串
         */
        private void dealStr(String encryptedStr) {
            System.out.println("处理字符串");
        }
        /**
         * <p>Description: 当前行数控制</p>
         */
        private void insert() {
            ++rowNumber;
        }

        /**
         * <p>Title: characters</p>
         * <p>Description: lastContents的获得</p>
         *
         * @param ch     字符数组
         * @param start  起始位
         * @param length 长度
         * @throws SAXException SAX异常
         * @see org.xml.sax.helpers.DefaultHandler#characters(char[], int, int)
         */
        public void characters(char[] ch, int start, int length) throws SAXException {
            lastContents += new String(ch, start, length);
        }
    }

}
