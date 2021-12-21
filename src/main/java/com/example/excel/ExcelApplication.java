package com.example.excel;

import cn.hutool.core.lang.Console;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.Excel03SaxReader;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

@SpringBootApplication
public class ExcelApplication {

    /*8
    合并excel
     */
    public static void main(String[] args) {
        //输入文件夹路径，扫描路劲下文件，读入内存(文件过大时可能存在内存溢出)
        System.out.println("请输入需合并的excel文件夹路径");
        List<Map<String,Object>> merge = new ArrayList<>();

        Scanner scanner = new Scanner(System.in);
        File directory = null;
        while (scanner.hasNextLine()){
            String line = scanner.nextLine();
            if(StrUtil.isNotBlank(line)){
                directory = new File(line);
                if(directory.isDirectory()){
                    File[] files = directory.listFiles();
                    for (int i = 0; i < files.length; i++) {
                        ExcelReader reader = ExcelUtil.getReader(files[i]);
                        merge.addAll(reader.readAll());
                    }

                    break;
                }else {
                    System.out.println("路径错误，请输入文件夹路径");
                }
            }
        }

        //写出
        File out = new File(directory,"merge.xlsx");
        BigExcelWriter writer= ExcelUtil.getBigWriter(out);
        // 一次性写出内容，使用默认样式
        writer.write(merge);
        // 关闭writer，释放内存
        writer.close();
    }

}
