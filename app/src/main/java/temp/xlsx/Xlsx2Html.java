package temp.xlsx;

import android.util.Log;
import android.util.Xml;

import org.xmlpull.v1.XmlPullParser;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Created by temp on 14-6-24.
 * 转化xlsx文件 为 html，
 */
public class Xlsx2Html {
    private static final String TAG = "Xlsx2Html";

    private static final String MAIN_FILE_NAME = "[Content_Types].xml";
    private static final String SHEET_RELATION_NAME = "xl/_rels/workbook.xml.rels";
    private final String xlsxPath;  //xlsx文件路径
    private String picCachePath;    //图片缓存目录

    public Xlsx2Html(String path) {
        xlsxPath = path;
    }


    private static void logd(String s) {
        Log.d(TAG, s);
    }

    /**
     * 转换xlsx文件为html格式，经验证，输出的string仅仅包括纯文本格式的txt<br>
     * xlsx, docx, pptx 本质上是一个zip文件，可以修改文件后缀名称，再解压即可看到内容
     */
    public String convert() throws Exception {
        StringBuilder html = new StringBuilder();
        html.append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>"
                + "<html><head> </head><body>");

        ArrayList<String> sheetList = new ArrayList<String>();
        String sharedStringFile = "", workbookStringFile = "";
        final ZipFile file = new ZipFile(new File(xlsxPath));  //打开供阅读的 ZIP 文件，由指定的 File 对象给出

        logd("1. parse [Content_Types].xml");

        ArrayList<String> t = setFileName(file, sheetList);
        sharedStringFile = t.get(0);
        workbookStringFile = t.get(1);
        t = null;

        logd("2. parsing workbook.xml, file :" + workbookStringFile);
        ArrayList<String> sheetNames = new ArrayList<String>();
        setSheetNames(file, workbookStringFile, sheetNames);

        logd("3. sort sheet name,");
        sortSheetName(file, sheetNames, sheetList);

        //解析sharedString.xml内容,将公用string添加到列表ls
        logd("4. parsing sharedString...");
        ArrayList<String> sharedStrings = new ArrayList<String>();//缓存部分cell的内容/String
        setSharedString(file, sharedStringFile, sharedStrings);

        logd("5. parsing every sheet Files, size : " + sheetList.size());
        boolean flat = false;
        for (String oneSheet : sheetList) {
            logd("---------start parse sheet file:  " + oneSheet);
            html.append("<p><p>" + sheetNames.remove(0) + "</p>");
            ZipEntry zipEntry = file.getEntry(oneSheet);
            InputStream inputStream = file.getInputStream(zipEntry);
            XmlPullParser xmlParserSheet = Xml.newPullParser();
            xmlParserSheet.setInput(inputStream, "UTF-8");
            int evtType = xmlParserSheet.getEventType();
            //firstly, find tag <mergeCell />
            ArrayList<String> mergeCells = new ArrayList<String>();
            while (evtType != XmlPullParser.END_DOCUMENT) {
                if (evtType != XmlPullParser.START_TAG ||
                        !"mergeCell".equalsIgnoreCase(xmlParserSheet.getName())) {
                    evtType = xmlParserSheet.next();
                    continue;
                }
                String mergeCell = xmlParserSheet.getAttributeValue(null, "ref");
                mergeCells.add(mergeCell);
                logd("find mergeCell ref : " + mergeCell); // eg: B4:D4
                evtType = xmlParserSheet.next();
            }
            logd("5.1 find mergeCell over! , start parse sheet.");

            ArrayList<String> mergePrefixs = new ArrayList<String>();
            for (String merge : mergeCells)  //find every prefix
                mergePrefixs.add(merge.substring(0, merge.indexOf(":")));    //eg:  B4:D4 --> B4
            // secondly, find every cell, parse it and append the html
            IOUtils.closeQuietly(inputStream);
            inputStream = file.getInputStream(zipEntry);
            xmlParserSheet = Xml.newPullParser();
            xmlParserSheet.setInput(inputStream, "UTF-8");
            evtType = xmlParserSheet.getEventType();
            String v = "";
            while (evtType != XmlPullParser.END_DOCUMENT) {
                switch (evtType) {
                    case XmlPullParser.START_TAG:
                        String tag = xmlParserSheet.getName();
                        String startTagHtml = tag2Html(xmlParserSheet, tag, true);
                        html.append(startTagHtml);
                        if ("c".equalsIgnoreCase(tag)) {                           // c  --> <td>
                            String str = xmlParserSheet.getAttributeValue(null, "t");
                            flat = (str != null);
                            str = xmlParserSheet.getAttributeValue(null, "r");//
                            if (mergePrefixs.contains(str)) {  //若mergeCells中包括 str
                                //通过 mergeCells设置此cell 的 rawspan, colspan
                                String forSpan = getCellSpan(str, mergeCells);
                                html.append("<td " + forSpan + ">");
                            } else {
                                html.append("<td>");
                            }
                        } else if ("v".equalsIgnoreCase(tag)) {
                            v = xmlParserSheet.nextText();
                            if (v != null) {
                                if (flat) {
                                    String value = sharedStrings.get(Integer.parseInt(v));
                                    html.append(value + " ");
                                } else {
                                    html.append(v + " ");
                                }
                            }
                        }
                        break;
                    case XmlPullParser.END_TAG:
                        String endTag = xmlParserSheet.getName();
                        String tagHtml = tag2Html(xmlParserSheet, endTag, false);
                        if (!html.toString().endsWith("<td>"))
                            html.append(tagHtml);
                        else {
                            html.delete(html.length() - 4, html.length());
                        }
                        break;
                }
                evtType = xmlParserSheet.next();
            }
            IOUtils.closeQuietly(inputStream);
            html.append("</p>");
            logd("---------end parse sheet file: " + oneSheet);
        }

        return html + "</body></html>";
    }

    /**
     * 根据[Content_Types].xml,找出相关文件路径: sheetList, sharedStringFile
     * @param file  表示xlsx文件
     * @param sheetList 要设置的sheet文件路径的列表
     * @return 一个ArrayList 第一个元素是sharedString.xml 第二个元素是workbook.xml的文件路径
     * */
    private ArrayList<String> setFileName(ZipFile file, ArrayList<String> sheetList) throws Exception {
        ArrayList<String> temp = new ArrayList<String>(2);  //size must be 2.
        temp.add("1"); temp.add("2");
        ZipEntry zipEntry = file.getEntry(MAIN_FILE_NAME);
        InputStream inputStream = file.getInputStream(zipEntry);
        XmlPullParser xmlPullParser = Xml.newPullParser();
        xmlPullParser.setInput(inputStream, "UTF-8");
        int evtType = xmlPullParser.getEventType();
        while (evtType != XmlPullParser.END_DOCUMENT) {
            switch (evtType) {
                case XmlPullParser.START_TAG:
                    String tag = xmlPullParser.getName();
                    String partName, contentType;
                    if ("Override".equalsIgnoreCase(tag)) {
                        contentType = xmlPullParser.getAttributeValue(null, "ContentType");
                        partName = xmlPullParser.getAttributeValue(null, "PartName");
                        if ("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
                                .equalsIgnoreCase(contentType)) { // xl/sharedStrings.xml
                            temp.set(0, partName.substring(1));
                            logd("find sharedString : " + partName);
                        } else if ("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                                .equalsIgnoreCase(contentType)) {
                            sheetList.add(partName.substring(1));
                            logd("find  sheet: " + partName);
                        } else if ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                                .equalsIgnoreCase(contentType)) {
                            temp.set(1, partName.substring(1));
                            logd("find workbook: " + partName);
                        }
                    }
                    break;
            }
            evtType = xmlPullParser.next();
        }
        IOUtils.closeQuietly(inputStream);
        return temp;
    }

    /**
     * 解析workbook.xml, 得出sheet个数和名称, 填充到sheetNames
     * @param file                  表示xlsx文件
     * @param workbookStringFile    workbook文件的路径
     * @param sheetNames            要设置的sheet名称列表
     */
    private void setSheetNames(ZipFile file, String workbookStringFile, ArrayList<String> sheetNames) throws Exception {
        ZipEntry zipEntry = file.getEntry(workbookStringFile);
        InputStream inputStream = file.getInputStream(zipEntry);
        XmlPullParser xmlPullParser = Xml.newPullParser();
        xmlPullParser.setInput(inputStream, "UTF-8");
        int evtType = xmlPullParser.getEventType();
        while (evtType != XmlPullParser.END_DOCUMENT) {
            switch (evtType) {
                case XmlPullParser.START_TAG:
                    String tag = xmlPullParser.getName();
                    if ("sheet".equalsIgnoreCase(tag)) {
                        String name = xmlPullParser.getAttributeValue(null, "name");
                        sheetNames.add(name);
                        logd("find sheetName: " + name);
                    }
                    break;
            }
            evtType = xmlPullParser.next();
        }
        IOUtils.closeQuietly(inputStream);
    }

    /**
     * 排序sheet name, 使之和 file name 列表对应,
     *
     * @param zipFile    表示xlsx文件
     * @param sheetNames sheet标签列表
     * @param filenames  待排序的sheet内容文件名的列表
     */
    private void sortSheetName(ZipFile zipFile, ArrayList<String> sheetNames, ArrayList<String> filenames) {
//        HashMap<String, String> name_file = new HashMap<String, String>();
        try {
            if (filenames.size() < 2) return;  // need no sort if only one sheet.
            ArrayList<Integer> ins = new ArrayList<Integer>();
            for (String fn : filenames) {       //从filenames中摘取出数字,存到ins列表
                int firstDigit = indexOfFirstDigit(fn);
                String digit = "";
                do {
                    digit += fn.charAt(firstDigit);
                    firstDigit++;
                } while (Character.isDigit(fn.charAt(firstDigit)));
                ins.add(Integer.parseInt(digit));
            }
            //sort ins, do same sort for filenames
            for (int i = 0; i < ins.size(); i++) {
                int min = ins.get(i);
                for (int j = i + 1; j < ins.size(); j++) {
                    if (min > ins.get(j)) {
                        min = ins.get(i);
                        ins.set(i, ins.get(j));
                        ins.set(j, min);
                        String temp = filenames.get(i);
                        filenames.set(i, filenames.get(j));
                        filenames.set(j, temp);
                    }
                }
            }
            // 目前以sheet文件名后面的数字大小来排序
            //:TODO maybe should sort according to rId, workbook.xml,and workbook.xml.rels.
//            ZipEntry sheetRelation = zipFile.getEntry(SHEET_RELATION_NAME);
//            InputStream is = zipFile.getInputStream(sheetRelation);
//            XmlPullParser xpp = Xml.newPullParser();
//            xpp.setInput(is, "utf-8");
//            int evtType = xpp.getEventType();
//            while (evtType != XmlPullParser.END_DOCUMENT) {
//                if (evtType == XmlPullParser.START_TAG) {
//                    String tag = xpp.getName();
//                    if (tag.equalsIgnoreCase("Relationship")) {
//                        String id = xpp.getAttributeValue(null, "Id"),
//                                type = xpp.getAttributeValue(null, "Type"),
//                                target = xpp.getAttributeValue(null, "Target");
//                        if (type.endsWith("worksheet"))     // must be sheet
//                            name_file.put(id,target);
//                    }
//                }
//                evtType = xpp.next();
//            }
//
//            // parse workbook.xml
//
        } catch (Exception e) {
            logd("sortSheetName failed. " + e);
        }
    }

    /**
     * 解析 sharedString文件， 获取共享的string列表，存到 list中
     */
    private void setSharedString(ZipFile file, String sharedStringFile, ArrayList<String> ls) throws Exception {
        ZipEntry zipEntry = file.getEntry(sharedStringFile); //返回指定名称的 ZIP 文件条目,打开文件xl/sharedStrings.xml
        InputStream inputStream = file.getInputStream(zipEntry);   // 返回输入流以读取指定 ZIP 文件条目的内容
        XmlPullParser xmlParser = Xml.newPullParser(); //Returns a new pull parser with namespace support.
        xmlParser.setInput(inputStream, "UTF-8");
        int evtType = xmlParser.getEventType();
        StringBuilder oneString = new StringBuilder();
        while (evtType != XmlPullParser.END_DOCUMENT) {// 以pull方式解析xml文件
            switch (evtType) {
                case XmlPullParser.START_TAG:
                    String tag = xmlParser.getName();
                    if ("t".equalsIgnoreCase(tag)) {
                        oneString.append(xmlParser.nextText());
                    }
                    break;
                case XmlPullParser.END_TAG:
                    if ("si".equalsIgnoreCase(xmlParser.getName())) { // one <si> </si> is a string
                        ls.add(oneString.toString());
                        oneString.setLength(0);
                    }
                    break;
                default:
                    break;
            }
            evtType = xmlParser.next();
        }
        IOUtils.closeQuietly(inputStream);
    }

    /**
     * 根据cellid和merageCells 来获取cell的span信息  返回span 描述
     */
    private String getCellSpan(String cellId, ArrayList<String> mCS) {
        String spanDes = "";
        //cellId must equal mCS's prefix.  eg  ( B2,   B2:C4 )
        String mergeCell = "";
        for (String s : mCS) {
            if (s.substring(0, s.indexOf(":")).equals(cellId))
                mergeCell = s;
        }
        String cells[] = mergeCell.split(":"); // get pre,suf, eg B2:C4 --> B2  C4
        int digitIndex0 = indexOfFirstDigit(cells[0]),
                digitIndex1 = indexOfFirstDigit(cells[1]);
        String strPre = cells[0].substring(0, digitIndex0).toUpperCase(),
                strSuf = cells[1].substring(0, digitIndex1).toUpperCase();
        int intPre = Integer.parseInt(cells[0].substring(digitIndex0)),
                intSuf = Integer.parseInt(cells[1].substring(digitIndex1));
        int colspan = 0;
        while (strPre.length() != strSuf.length()) {// must be : pre < suf.
            colspan += 26 * (strSuf.charAt(0) - 'A' + 1);
            strSuf = strSuf.substring(1);
        }
        //now strPre.length() == strSuf.length()  AB  BD  , result is (B-A+1)*26 + (D-B+1)
        while (strPre.length() > 0) {
            int len1 = strPre.length(), len2 = strSuf.length();
            colspan += strSuf.charAt(len2 - 1) - strPre.charAt(len1 - 1) + 1;
            strPre = strPre.substring(0, len1 - 1);
            strSuf = strSuf.substring(0, len2 - 1);
        }
        if (colspan > 0)
            spanDes += " COLSPAN=" + colspan;
        int rawSpan = intSuf - intPre + 1;
        if (colspan > 0)
            spanDes += " ROWSPAN=" + rawSpan;
        return spanDes;
    }

    /**
     * 返回string的第一个数字的索引
     */
    private int indexOfFirstDigit(String str) {
        for (int i = 0; i < str.length(); i++) {
            if (Character.isDigit(str.charAt(i)))
                return i;
        }
        return -1;
    }

    private String tag2Html(XmlPullParser xpp, String tag, boolean start) {
        String h = "";
        if ("sheetData".equalsIgnoreCase(tag)) {
            h = start ? "<table class=excelDefaults border=\"1\">" : "</table>";
        } else if ("row".equalsIgnoreCase(tag)) {
            String height = "-1";
            if (start) {
                height = xpp.getAttributeValue(null, "ht");//行的高度
                try {
                    height = (int) (Integer.parseInt(height) / 3.0 * 4) + "";
                } catch (NumberFormatException e) {
                }
            }
            if ("-1".equalsIgnoreCase(height)) {
                h = start ? "<tr>" : "</tr>";
            } else {
                h = /*start must true.*/"<tr height=" + height + ">";
            }
        } else if ("c".equalsIgnoreCase(tag)) {
            h = start ? "" : "</td>"; // if start true, do it in convert().
        }
        return h;
    }

}
