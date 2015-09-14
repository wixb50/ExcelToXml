package com.xie.excelxml;

import java.io.File;
import java.io.FilenameFilter;

/**
 * Created by int on 15-9-7.
 */
//定义扩展名过滤器
class FileExtFilter implements FilenameFilter {
    String extent;

    FileExtFilter(String extent) {
        this.extent = extent;
    }

    @Override
    public boolean accept(File dir, String name) {
        // TODO Auto-generated method stub
        return name.endsWith("." + extent);
    }
}
