package com.company;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class Main {

    private static File lostFileTxt;
    private static String destDirInput;
    private static String sourceDirInput;

    public static void main(String[] args) {
        System.out.println("程序开始执行");
        destDirInput = "";
        sourceDirInput = "";

        System.out.println("请输入需要填写统计表的路径:");
        Scanner s = new Scanner(System.in);
        destDirInput = s.nextLine();
        System.out.println("请输入导出表的路径:");
        sourceDirInput = s.nextLine();


        int destFilesCnt = 0;
        int completeFilesCnt = 0;
        int currentCompleteFileCnt = 0;
        ArrayList<String> lostFileNameList = new ArrayList();
        //"/Users/jonesleborn/Desktop/XYReadExcel/dest"
        File destDir = new File(destDirInput);
        if (destDir.exists()) {
            File[] destFiles = destDir.listFiles();
            destFilesCnt = destFiles.length;
            initLostFileTxt();
            for (int i = 0; i < destFiles.length; i++) {
                File destFile = destFiles[i];
                if (destFile.getName().endsWith(".xlsx")) {
                    String footContent = ExcelUtil.deleteRow(destFile);
                    System.out.println("读取到的footContent-->" + footContent);
                    // 读取/dest统计表中的数据，获取图幅编码
                    String picCode = ExcelUtil.readPicCode(destFile);

                    // 读取/source中的文件列表，根据picCode拿到导出数据表
                    //"/Users/jonesleborn/Desktop/XYReadExcel/source"
                    File sourceDir = new File(sourceDirInput);
                    if (sourceDir.exists()) {
                        // 列出所有文件
                        File[] sourceFiles = sourceDir.listFiles();
                        for (int j = 0; j < sourceFiles.length; j++) {
                            File sourceFile = sourceFiles[j];
                            if (sourceFile.getName().endsWith(".xlsx")) {
                                if (sourceFile.getName().contains(picCode)) {
                                    // 读取sourceFile中的内容
                                    ArrayList<SourceFileBean> list = ExcelUtil.ReadSourceFile(sourceFile);
                                    // 将导出数据表，写入到统计表中
                                    ExcelUtil.writeExcel(list, destFile.getAbsolutePath(), footContent);
                                    completeFilesCnt++;
                                    currentCompleteFileCnt++;
                                    System.out.println("写入文件" + destFile.getName() + "成功! " + "当前进度" + currentCompleteFileCnt + "/" + destFilesCnt);
                                    break;
                                }

                                if (j == sourceFiles.length - 1) {
                                    currentCompleteFileCnt++;
                                    System.out.println("====找不到文件" + picCode + ".xlsx，请手工确认====" + "当前进度" + currentCompleteFileCnt + "/" + destFilesCnt);
                                    lostFileNameList.add(picCode + ".xlsx");
                                }
                            }
                        }
                    } else {
                        System.out.println("sourceDir路径没有创建");
                    }
                }
            }
        } else {
            System.out.println("destDir路径没有创建");
        }
        System.out.println("程序执行完毕，已完成文件：" + completeFilesCnt + "|缺失文件：" + (destFilesCnt - completeFilesCnt));
        System.out.println("丢失的文件已经写入" + lostFileTxt.getAbsolutePath());

        try {
            BufferedWriter out = new BufferedWriter(new FileWriter(lostFileTxt, true));
            for (int i = 0; i < lostFileNameList.size(); i++) {
                out.write(lostFileNameList.get(i));
                out.newLine();
            }
            out.flush();
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void initLostFileTxt() {
        // 生成丢失文件.txt
        lostFileTxt = new File(destDirInput + "/lost.txt");
        try {
            if (lostFileTxt.exists()) {
                lostFileTxt.createNewFile();
            }
        } catch (Exception e) {
            System.out.println("创建lost.txt文件失败");
            e.printStackTrace();
        }
    }
}
