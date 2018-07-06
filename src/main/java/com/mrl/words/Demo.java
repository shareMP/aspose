package com.mrl.words;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo
{
    public static void main(String[] args) {
        
        
      /*  String tempPath = ClassPathUtil.getDeployWarPath() + "njzwfw/xmz/report/template/市级部门办件量统计表.xlsx";
        String path = ClassPathUtil.getDeployWarPath() + "njzwfw/xmz/report/export/";
        JSONObject json = JSON.parseObject(params);
        JSONObject obj = (JSONObject) json.get("params");
        String year = obj.getString("year");
        String month = obj.getString("month");

        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
        list = iTJProjectCityService.getTJProjectCityList(year, month);
        String fileName = "市级部门办件量统计表_" + System.currentTimeMillis() + ".xlsx";

        exportExcel3("city", tempPath, path, fileName, list);
        JSONObject dataJson = new JSONObject();
        dataJson.put("url", "njzwfw/xmz/report/export/" + fileName);
        return JsonUtils.zwdtRestReturn("1", "", dataJson);*/
     
        
        //模板
        List<Map<String, Object>> map = new ArrayList<Map<String, Object>>();
        Map<String,Object> map1 = new HashMap<>();
        map1.put("c1", "v1");
        map1.put("c2", "v2");
        
        Map<String,Object> map2 = new HashMap<>();
        map2.put("c1", "v3");
        map2.put("c2", "v4");
        map.add(map2);
        map.add(map1);
        
        File newFile = createNewFile("E:/ccc.xlsx", "E:/", "ggg.xlsx");
        InputStream is = null;
        XSSFWorkbook workbook = null;
        XSSFSheet sheet = null;

        try {
            is = new FileInputStream(newFile);// 将excel文件转为输入流
            workbook = new XSSFWorkbook(is);// 创建个workbook，
            // 获取第一个sheet
            sheet = (XSSFSheet) workbook.getSheetAt(0);
        }
        catch (Exception e1) {
            e1.printStackTrace();
        }

        if (sheet != null) {
            try {
                // 写数据
                FileOutputStream fos = new FileOutputStream(newFile);

                XSSFRow row = sheet.getRow(0);
                if (row == null) {
                    row = sheet.createRow(0);
                }
                //绘制第八列,年份
                XSSFCell cell = row.getCell(7);
                if (cell == null) {
                    cell = row.createCell(7);
                }
                
                XSSFRow row2 = sheet.getRow(1);
                if (row2 == null) {
                    row2 = sheet.createRow(1);
                }
               

                XSSFCell cell2 = row2.getCell(2);
                if (cell2 == null) {
                    cell2 = row.createCell(2);
                }
                cell2.setCellValue("300");
                XSSFCellStyle style2 = workbook.createCellStyle();
                style2.setAlignment(CellStyle.ALIGN_CENTER);
                Font f = workbook.createFont();
                f.setColor(Font.COLOR_RED);
                style2.setFont(f);
                cell2.setCellStyle(style2);
                
                
                XSSFCell cel3 = row2.getCell(4);
                if (cel3 == null) {
                    cel3 = row.createCell(4);
                }
                cel3.setCellValue("49.2%");
//               
                style2.setFont(f);
                cel3.setCellStyle(style2);
                
                XSSFCell cell4 = row2.getCell(8);
                if (cell4 == null) {
                    cell4 = row.createCell(4);
                }
                cell4.setCellValue("999");
               
                style2.setFont(f);
                cell4.setCellStyle(style2);
                
                
                
                
                XSSFCellStyle style;
                XSSFColor color;
                for (int i = 0; i < map.size(); i++) {
                    row = sheet.createRow(i + 5); //从第五行开始

                    style = workbook.createCellStyle();
                    color = new XSSFColor(new java.awt.Color(255, 255, 255));
                    
                    XSSFFont createFont = workbook.createFont();
                    createFont.setFontHeightInPoints((short)10);
                    
                    createFont.setFontName("宋体");
                    
                    style.setFont(createFont);
                    
                   createRowAndCell(String.valueOf(i+1), row, cell, 0, style, color);

                    //全程网办部门自建系统  【颜色 255,242,204】
                    //style = workbook.createCellStyle();
                    color = new XSSFColor(new java.awt.Color(255, 242, 204));
                    createRowAndCell(map.get(i).get("c1"), row, cell, 1, style, color);
                    createRowAndCell(map.get(i).get("c2"), row, cell, 2, style, color);
                    createRowAndCell(map.get(i).get("c1"), row, cell, 3, style, color);
                    createRowAndCell("测试测试测试吃撒打算打算打是=====sadjasjld 查sad啊实打实的是sss", row, cell, 4, style, color);
//                    createRowAndCell(map.get(i).get("instyear_zg_allnet"), row, cell, 3, style, color);
//                    createRowAndCell(map.get(i).get("zg_allnet_per"), row, cell, 4, style, color);

                  /*  createRowAndCell(map.get(i).get("instmonth_zg_window"), row, cell, 11, style, color);
                    createRowAndCell(map.get(i).get("lastmonth_zg_window"), row, cell, 12, style, color);
                    createRowAndCell(map.get(i).get("instyear_zg_window"), row, cell, 13, style, color);
                    createRowAndCell(map.get(i).get("zg_window_per"), row, cell, 14, style, color);

                    //全程网办部门自建系统 【颜色 252,228,214】
                    style = workbook.createCellStyle();
                    color = new XSSFColor(new java.awt.Color(252, 228, 214));
                    createRowAndCell(map.get(i).get("instmonth_zj_allnet"), row, cell, 5, style, color);
                    createRowAndCell(map.get(i).get("lastmonth_zj_allnet"), row, cell, 6, style, color);
                    createRowAndCell(map.get(i).get("instyear_zj_allnet"), row, cell, 7, style, color);
                    createRowAndCell(map.get(i).get("zg_allnet_per"), row, cell, 8, style, color);

                    createRowAndCell(map.get(i).get("instmonth_zj_window"), row, cell, 15, style, color);
                    createRowAndCell(map.get(i).get("lastmonth_zj_window"), row, cell, 16, style, color);
                    createRowAndCell(map.get(i).get("instyear_zj_window"), row, cell, 17, style, color);
                    createRowAndCell(map.get(i).get("zj_window_per"), row, cell, 18, style, color);
                    //合计、百分比  【颜色 226, 239,218】
                    style = workbook.createCellStyle();
                    color = new XSSFColor(new java.awt.Color(226, 239, 218));
                    createRowAndCell(map.get(i).get("totalyear_proj_allnet"), row, cell, 9, style, color);
                    createRowAndCell(map.get(i).get("total_allnet_per"), row, cell, 10, style, color);
                    //合计、百分比  【颜色 217, 225,242】
                    style = workbook.createCellStyle();
                    color = new XSSFColor(new java.awt.Color(217, 225, 242));
                    createRowAndCell(map.get(i).get("totalyear_proj_window"), row, cell, 19, style, color);
                    createRowAndCell(map.get(i).get("total_window_per"), row, cell, 20, style, color);

                    //总数 【颜色 237, 125,49】
                    style = workbook.createCellStyle();
                    color = new XSSFColor(new java.awt.Color(237, 125, 49));
                    createRowAndCell(map.get(i).get("totalyear_proj"), row, cell, 21, style, color);*/
                }
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
            catch (Exception e) {
                e.printStackTrace();
            }
            finally {
                try {
                    if (null != is) {
                        is.close();
                    }
                }
                catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        
    }
    
    /**
     * 读取excel模板，并复制到新文件中供写入和下载
     *  [功能详细描述]
     *  @param tempPath
     *  @param rPath
     *  @param newFileName
     *  @return    
     * @exception/throws [违例类型] [违例说明]
     * @see [类、类#方法、类#成员]
     */
    public static File createNewFile(String tempPath, String rPath, String newFileName) {
        // 读取模板，并赋值到新文件************************************************************
        // 文件模板路径
        String path = (tempPath);
        File file = new File(path);
        // 保存文件的路径
        String realPath = rPath;
        // 新的文件名
        //String newFileName = fileName + "_" + System.currentTimeMillis() + ".xlsx";
        // 判断路径是否存在
        File dir = new File(realPath);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        // 写入到新的excel
        File newFile = new File(realPath, newFileName);
        try {
            newFile.createNewFile();
            // 复制模板到新文件
            fileChannelCopy(file, newFile);
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        return newFile;
    }

    /**
     * 下载成功后删除
     * 
     * @param files
     */
    private static void deleteFile(File... files) {
        for (File file : files) {
            if (file.exists()) {
                file.delete();
            }
        }
    }
    
    public static void fileChannelCopy(File s, File t) {
        try {
            InputStream in = null;
            OutputStream out = null;
            try {
                in = new BufferedInputStream(new FileInputStream(s), 1024);
                out = new BufferedOutputStream(new FileOutputStream(t), 1024);
                byte[] buffer = new byte[1024];
                int len;
                while ((len = in.read(buffer)) != -1) {
                    out.write(buffer, 0, len);
                }
            }
            finally {
                if (null != in) {
                    in.close();
                }
                if (null != out) {
                    out.close();
                }
            }
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    /**
     *根据当前row行，来创建index标记的列数,并赋值数据
     */
    private static void createRowAndCell(Object obj, XSSFRow row, XSSFCell cell, int index, XSSFCellStyle style,
            XSSFColor color) {
        cell = row.getCell(index);
        if (cell == null) {
            cell = row.createCell(index);
        }

        if (obj != null)
            cell.setCellValue(obj.toString());
        else
            cell.setCellValue("");

        if (style != null) {
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中 
            //style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框    
            //style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框    
            //style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
            //style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框 
            //设置背景颜色
           // style.setFillForegroundColor(color);
            //style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
           
            
            cell.setCellStyle(style);
        }
    }
}
