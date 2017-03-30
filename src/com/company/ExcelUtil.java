package com.company;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class ExcelUtil {
    //默认单元格内容为数字时格式
    private static DecimalFormat df = new DecimalFormat("0");
    // 默认单元格格式化日期字符串
    private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    // 格式化数字
    private static DecimalFormat nf = new DecimalFormat("0.000");

    /**
     * 读取图幅号
     * <p>
     * 读取
     *
     * @param file dest文件
     * @return
     */
    public static String readPicCode(File file) {
        String value = "";
        StringBuilder sb = new StringBuilder();
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            for (int i = sheet.getFirstRowNum(), rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows(); i++) {
                row = sheet.getRow(i);
                if (row == null) {
                    //当读取行为空时
                    continue;
                } else {
                    rowCount++;
                }
                for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                    if (j < 0) break;
                    cell = row.getCell(j);
                    if (cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                        //当该单元格为空
                        if (j != row.getLastCellNum()) {//判断是否是该行中最后一个单元格
                        }
                        continue;
                    }
                    switch (cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                            if (i == 3 && j == 0) {
                                String picCode = cell.getStringCellValue();
                                String code = picCode.substring(picCode.lastIndexOf("：") + 1);
                                sb.append(code);
                            }
                            break;
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return sb.toString();
    }

    public static ArrayList<SourceFileBean> ReadSourceFile(File file) {
        ArrayList<SourceFileBean> listBean = new ArrayList<>();
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            Object value;
            for (int i = sheet.getFirstRowNum(), rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows(); i++) {
                row = sheet.getRow(i);
                if (row == null) {
                    //当读取行为空时
                    continue;
                } else {
                    rowCount++;
                }
                SourceFileBean bean = new SourceFileBean();
                for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                    if (j < 0) break;
                    cell = row.getCell(j);
                    if (cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                        //当该单元格为空
                        if (j != row.getLastCellNum()) {//判断是否是该行中最后一个单元格
                        }
                        continue;
                    }
                    switch (cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                            if (j == 9) {
                                // 备注
                                bean.setRemark(cell.getStringCellValue());
                            }
                            //System.out.println(i + "行" + j + " 列 is String type" + cell.getStringCellValue());
                            value = cell.getStringCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                value = df.format(cell.getNumericCellValue());
                            } else if ("General".equals(cell.getCellStyle()
                                    .getDataFormatString())) {
                                value = nf.format(cell.getNumericCellValue());
                            } else {
                                value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                        .getNumericCellValue()));
                            }
                            /*
                            System.out.println(i + "行" + j
                                    + " 列 is Number type ; DateFormt:"
                                    + value.toString());
                                    */
                            if (j == 0) {
                                // 序号
                                bean.setOrder(value.toString());
                            } else if (j == 1) {
                                // 点号
                                bean.setPoint(value.toString());
                            } else if (j == 2) {
                                // 影响X
                                bean.setX1(value.toString());
                            } else if (j == 3) {
                                // 影像Y
                                bean.setY1(value.toString());
                            } else if (j == 4) {
                                // x2
                                bean.setX2(value.toString());
                            } else if (j == 5) {
                                // y2
                                bean.setY2(value.toString());
                            } else if (j == 6) {
                                // dx
                                bean.setDx(value.toString());
                            } else if (j == 7) {
                                // dy
                                bean.setDy(value.toString());
                            } else if (j == 8) {
                                bean.setDs(value.toString());
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            //System.out.println(i + "行" + j + " 列 is Boolean type");
                            //value = Boolean.valueOf(cell.getBooleanCellValue());
                            break;
                        case XSSFCell.CELL_TYPE_BLANK:
                            //System.out.println(i + "行" + j + " 列 is Blank type");
                            //value = "";
                            break;
                        default:
                            //System.out.println(i + "行" + j + " 列 is default type");
                            //value = cell.toString();
                    }
                }
                if (bean.getOrder() != null) {
                    listBean.add(bean);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println(file.getName() + "--读取" + file.getName() + "有问题");
        }
        return listBean;
    }


    public static String deleteRow(File targetFile) {
        String footPrint = "";
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(targetFile));
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            int physicalRows = sheet.getPhysicalNumberOfRows();
            if (targetFile.getName().contains("dom百色F48G003070")) {
                System.out.println("asd");
            }
            for (int i = sheet.getFirstRowNum(), rowCount = 0; rowCount < physicalRows; i++) {
                row = sheet.getRow(i);
                if (row == null) {
                    //当读取行为空时
                    continue;
                } else {
                    rowCount++;
                }
                if (rowCount >= 9) {
                    // 删除行
                    sheet.removeRow(row);
                    for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                        if (j < 0) break;
                        cell = row.getCell(j);
                        if (cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                            //当该单元格为空
                            if (j != row.getLastCellNum()) {//判断是否是该行中最后一个单元格
                            }
                            continue;
                        }
                        switch (cell.getCellType()) {
                            case XSSFCell.CELL_TYPE_STRING:
                                if (cell.getStringCellValue().contains("检查者：")) {
                                    footPrint = cell.getStringCellValue();
                                }
                                break;
                        }
                    }
                }

            }

            //int regionsCnt = sheet.getNumMergedRegions();
            /*
            for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {
                // Delete the region
                //System.out.println("regionsCnt-->" + regionsCnt);
                sheet.removeMergedRegion(i);
            }*/

            /*
            for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
                // Delete the region
                //System.out.println("regionsCnt-->" + regionsCnt);

                if (i == sheet.getNumMergedRegions() - 1 || i == sheet.getNumMergedRegions() - 2) {
                    sheet.removeMergedRegion(i);
                }
                if (i == ) {
                    sheet.removeMergedRegion(i);
                }
            }*/

            sheet.removeMergedRegion(sheet.getNumMergedRegions() - 1);

            FileOutputStream os = new FileOutputStream(targetFile.getAbsolutePath());
            wb.write(os);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println(targetFile.getName() + "删除空行方法有问题，出错文件" + targetFile.getName());
        }
        return footPrint;
    }

    public static void writeExcel(ArrayList<SourceFileBean> result, String path, String footContent) {
        if (result == null) {
            return;
        }
        File targetFile = new File(path);

        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(targetFile));
            XSSFSheet sheet = wb.getSheet("Sheet1");
            XSSFDataFormat dataFormat = wb.createDataFormat();
            int rowIndex = sheet.getLastRowNum();
            XSSFCellStyle cellStyle = wb.createCellStyle();
            cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
            cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
            cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
            cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);

            for (int i = 0; i < result.size(); i++) {
                // 插入最后一行
                rowIndex = rowIndex + 1;
                XSSFRow row = sheet.createRow(rowIndex);
                XSSFCell cell;
                // 根据sourceBean有多少列来
                for (int j = 0; j < 10; j++) {
                    cell = row.createCell(j);
                    cell.setCellStyle(cellStyle);
                    cellStyle.setDataFormat(dataFormat.getFormat("General"));
                    /*
                    if (j == 0) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getOrder()));
                    } else if (j == 1) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getPoint()));
                    } else if (j == 2) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getX1()));
                    } else if (j == 3) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getY1()));
                    } else if (j == 4) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getX2()));
                    } else if (j == 5) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getY2()));
                    } else if (j == 6) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getDx()));
                    } else if (j == 7) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getDy()));
                    } else if (j == 8) {
                        cell.setCellValue(Double.parseDouble(result.get(i).getDs()));
                    } else if (j == 9) {
                        cell.setCellValue(result.get(i).getRemark());
                    }
                    */
                }
                cell = sheet.getRow(rowIndex).getCell(0);
                cell.setCellValue(Double.parseDouble(result.get(i).getOrder()));

                cell = sheet.getRow(rowIndex).getCell(1);
                cell.setCellValue(Double.parseDouble(result.get(i).getPoint()));

                cell = sheet.getRow(rowIndex).getCell(2);
                cell.setCellValue(Double.parseDouble(result.get(i).getX1()));

                cell = sheet.getRow(rowIndex).getCell(3);
                cell.setCellValue(Double.parseDouble(result.get(i).getY1()));

                cell = sheet.getRow(rowIndex).getCell(4);
                cell.setCellValue(Double.parseDouble(result.get(i).getX2()));

                cell = sheet.getRow(rowIndex).getCell(5);
                cell.setCellValue(Double.parseDouble(result.get(i).getY2()));

                cell = sheet.getRow(rowIndex).getCell(6);
                cell.setCellValue(Double.parseDouble(result.get(i).getDx()));

                cell = sheet.getRow(rowIndex).getCell(7);
                cell.setCellValue(Double.parseDouble(result.get(i).getDy()));

                cell = sheet.getRow(rowIndex).getCell(8);
                cell.setCellValue(Double.parseDouble(result.get(i).getDs()));

                cell = sheet.getRow(rowIndex).getCell(9);
                cell.setCellValue(result.get(i).getRemark());
            }

            // 写入foot
            rowIndex++;
            XSSFRow row = sheet.createRow(rowIndex);
            // 先创建好10个空格，并且设置了边框
            for (int k = 0; k < 10; k++) {
                XSSFCell cell = row.createCell(k);
                cell.setCellStyle(cellStyle);
                cellStyle.setDataFormat(dataFormat.getFormat("General"));
                if (k == 0) {
                    cell.setCellValue(footContent);
                }
            }
            // 在将10个空格合并成一个单元格,写入footer
            //CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 9);
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 9));
            //XSSFRow footRow = sheet.getRow(rowIndex);

            //XSSFCell cell = footRow.createCell(0);

            //cell.setCellStyle(cellStyle);
            //cell.setCellValue(footContent);
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            try {
                wb.write(os);
            } catch (IOException e) {
                e.printStackTrace();
            }
            byte[] content = os.toByteArray();
            File file = new File(path);//Excel文件生成后存储的位置。
            OutputStream fos = null;
            try {
                fos = new FileOutputStream(file);
                fos.write(content);
                os.close();
                fos.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            System.out.println("写入" + targetFile.getName() + "出错");
            e.printStackTrace();
        }

    }

    public static DecimalFormat getDf() {
        return df;
    }

    public static void setDf(DecimalFormat df) {
        ExcelUtil.df = df;
    }

    public static SimpleDateFormat getSdf() {
        return sdf;
    }

    public static void setSdf(SimpleDateFormat sdf) {
        ExcelUtil.sdf = sdf;
    }

    public static DecimalFormat getNf() {
        return nf;
    }

    public static void setNf(DecimalFormat nf) {
        ExcelUtil.nf = nf;
    }


}
