package com.xie.excelxml;

import java.io.File;
import java.util.Scanner;

/**
 * com.xie.excelxml.ExcelXML.java
 * eleganceTse
 * 2015-9-1 上午08:14:59
 */
public class ExcelToXml {

    /**
     * 2010-6-3 上午08:14:59
     *
     * @param args
     */
    public static void main(String[] args) {
        Excel sourceExcel = null;
        Scanner scanner = new Scanner(System.in);
        String xmlFileName = null;
        try {
            System.out.println("************************\n");
            System.out.println("请输入你的选择：\n1.转换一个excel到xml;\n2.转换当前所有excel到xml.\n");
            System.out.println("************************\n");
            int choose = scanner.nextInt();
            switch (choose) {
                case 1:
                    System.out.println("输入文件路径（带文件名），相对绝对都可以：\n");
                    String filePath = scanner.next();
                    sourceExcel = new Excel(filePath);
                    //设置保存xml文件名称
                    String tempName = new File(filePath).getName();
                    xmlFileName = tempName.substring(0, tempName.indexOf("."));
                    //调用函数转换
                    sourceExcel.allSheetToXml(xmlFileName);
                    System.out.println("\n******转换成功！！！******");
                    break;
                case 2:
                    //筛选目录下的excel文件
                    FileExtFilter fileExtFilter = new FileExtFilter("xls");
                    File fileDir = new File(".");
                    String[] fileList = fileDir.list(fileExtFilter);
                    for (int i = 0; i < fileList.length; i++) {
                        System.out.println(fileList[i] + "-->>");//调用函数转换
                        sourceExcel = new Excel(fileList[i]);
                        //设置保存xml文件名称
                        String tempName1 = new File(fileList[i]).getName();
                        xmlFileName = tempName1.substring(0, tempName1.indexOf("."));
                        //调用函数转换
                        sourceExcel.allSheetToXml(xmlFileName);
                    }
                    System.out.println("\n******转换成功！！！******");
                    break;
                default:
                    System.out.println("输入错误");
                    break;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}