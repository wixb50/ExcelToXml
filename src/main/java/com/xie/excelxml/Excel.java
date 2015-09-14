package com.xie.excelxml;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.DocumentType;
import org.dom4j.Element;
import org.dom4j.dom.DOMDocumentType;
import org.dom4j.dtd.ElementDecl;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

/**
 * @author wixb 设计一个类，用以封装操作excel表的各种操作
 */
public class Excel {
    // 当前excel的路径
    public String filePath;
    // 当前excel工作簿，唯一标识一个excel文件
    public Workbook workBook;
    // excel内的表格，可有多个
    public Sheet sheet;
    // 指定表格sheet的行
    public Row row;
    // 指定行的某个单元格
    public Cell cell;

    /**
     * 构造函数，初始化值
     *
     * @param excel
     * @throws Exception
     */
    public Excel(String excel) throws Exception {
        this.filePath = excel;
        InputStream in = new FileInputStream(excel);
        this.workBook = WorkbookFactory.create(in);
        // 默认指向第一个表，初始化
        this.sheet = this.workBook.getSheetAt(0);
        if ((this.row = this.sheet.getRow(0)) == null)
            this.row = this.sheet.createRow(0);
        if ((this.cell = this.row.getCell(0)) == null)
            this.cell = this.row.createCell(0);
    }


    /**
     * 获得指定位置的值，返回object，可为数值，字符串，布尔类型，null类型
     *
     * @param column
     * @return
     */
    public Object getCellValueObject(int rowNum, int column) {
        // 定义返回的数组
        Object tempObject = null;
        row = sheet.getRow(rowNum);
        cell = row.getCell(column);
        // 判断值类型
        switch (cell.getCellType()) {
            // 字符串类型
            case Cell.CELL_TYPE_STRING:
                tempObject = cell.getRichStringCellValue().getString();
                //System.out.println(cell.getRichStringCellValue().getString());
                break;
            // 数值类型
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    tempObject = cell.getDateCellValue();
                    //System.out.println(cell.getDateCellValue());
                } else {
                    tempObject = cell.getNumericCellValue();
                    //System.out.println(cell.getNumericCellValue());
                }
                break;
            // 布尔类型
            case Cell.CELL_TYPE_BOOLEAN:
                tempObject = cell.getBooleanCellValue();
                //System.out.println(cell.getBooleanCellValue());
                break;
            // 数学公式类型
            case Cell.CELL_TYPE_FORMULA:
                tempObject = cell.getCellFormula();
                //System.out.println(cell.getCellFormula());
                break;
            default:
                System.out.println();
        }
        return tempObject;
    }

    /**
     * 将excel内的内容读取到xml文件中，并添加dtd验证
     *
     * @param xmlFile
     * @param sheetNum
     * @return 1代表成功，0失败，-1超过最大sheet
     */
    public int excelToXml(String xmlFile, int sheetNum) {
        if (sheetNum >= workBook.getNumberOfSheets())
            return -1;
        else
            sheet = workBook.getSheetAt(sheetNum);
        xmlFile = xmlFile + ".xml";
        try {
            Document document = DocumentHelper.createDocument();
            //使用sheet名称命名跟节点
            String rootName = sheet.getSheetName().replaceAll(" ", "");
            Element root = document.addElement(rootName);
            //添加dtd文件说明
            DocumentType documentType = new DOMDocumentType();
            documentType.setElementName(rootName);
            List<ElementDecl> declList = new ArrayList<>();
            declList.add(new ElementDecl(rootName, "(row*)"));
            //判断sheet是否为空,为空则不执行任何操作
            if (sheet.getRow(0) == null)
                return 1;
            //遍历sheet第一行,获取元素名称
            row = sheet.getRow(0);
            String rowString = null;
            List<String> pcdataList = new ArrayList<>();
            for (int y = 0; y < row.getPhysicalNumberOfCells(); y++) {
                Object object = this.getCellValueObject(0, y);
                if (rowString != null)
                    rowString += "|" + object.toString();
                else
                    rowString = object.toString();
                pcdataList.add(object.toString());
            }
            //设置行节点
            declList.add(new ElementDecl("row", "(" + rowString + ")*"));
            //遍历list设置行的下级节点
            for (String tmp : pcdataList) {
                declList.add(new ElementDecl(tmp, "(#PCDATA)"));
            }
            documentType.setInternalDeclarations(declList);
            //遍历读写excel数据到xml中
            for (int x = 1; x < sheet.getLastRowNum(); x++) {
                row = sheet.getRow(x);
                Element rowElement = root.addElement("row");
                for (int y = 0; y < row.getPhysicalNumberOfCells(); y++) {
                    //cell = row.getCell(y);
                    Object object = this.getCellValueObject(x, y);
                    if (object != null) {
                        //将sheet第一行的行首元素当作元素名称
                        String pcdataString = pcdataList.get(y);
                        Element element = rowElement.addElement(pcdataString);
                        //Element element = rowElement.addElement("name");
                        element.setText(object.toString());
                    }
                }
            }
            //写入文件和dtd
            document.setDocType(documentType);
            this.docToXmlFile(document, xmlFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return 1;
    }

    /**
     * 转换excel中的所有sheet
     *
     * @param xmlFile
     * @return
     */
    public boolean allSheetToXml(String xmlFile) {
        for (int i = 0; i < workBook.getNumberOfSheets(); i++) {
            if (excelToXml(xmlFile + "_" + i, i) == 0) {
                System.out.println("转换出错！");
                return false;
            }
        }
        return true;
    }

    /*下面是xml封装的函数*/

    /**
     * 将document写入xml文件中
     *
     * @param document
     * @param xmlFile
     * @return
     */
    public boolean docToXmlFile(Document document, String xmlFile) {
        try {
            // 排版缩进的格式
            OutputFormat format = OutputFormat.createPrettyPrint();
            // 设置编码,可按需设置
            format.setEncoding("gb2312");
            // 创建XMLWriter对象,指定了写出文件及编码格式
            XMLWriter writer = new XMLWriter(new OutputStreamWriter(
                    new FileOutputStream(new File(xmlFile)), "gb2312"), format);
            // 写入
            writer.write(document);
            // 立即写入
            writer.flush();
            // 关闭操作
            writer.close();
            System.out.println("输出xml文件到：" + xmlFile);
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }
}
