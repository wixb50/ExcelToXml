package com.xie.excelxml;

import java.io.File;
import java.io.FileWriter;
import java.io.PrintWriter;

/**
 * Created by int on 15-10-9.
 */
public class Util {
    /**
     * 将错误信息写入log文件
     *
     * @param logFile
     * @param msg
     * @return
     */
    public static boolean msgToLog(String logFile, String msg) {
        try {
            File myFile = new File(logFile);
            if (!myFile.exists())
                myFile.createNewFile();
            FileWriter resultFile = new FileWriter(myFile, true);
            //把该对象包装进PrinterWriter对象
            PrintWriter myNewFile = new PrintWriter(resultFile);
            myNewFile.println(msg);
            resultFile.close();   //关闭文件写入流
        } catch (Exception e) {
            System.out.println("无法创建新文件！");
            e.printStackTrace();
        }
        return true;
    }
}
