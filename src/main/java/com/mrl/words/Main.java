package com.mrl.words;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Test;

import com.aspose.words.CellVerticalAlignment;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.HeightRule;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.SaveFormat;
import com.aspose.words.Table;

public class Main
{
    public static void main(String[] args) throws Exception {

        // 注册码
        // String licenseName = Main.class.getResource("/").getPath() +
        // "license.xml";
        // String licenseName =
        // "D:/个人工作空间/study/words/src/main/resources/license.xml";
       /* InputStream inputLincese = Main.class.getClassLoader().getResourceAsStream("license.xml");
        License license = new License();
        license.setLicense(inputLincese);

        InputStream input = Main.class.getClassLoader().getResourceAsStream("模板.docx");
        Document document = new Document(input);
        DocManager docManager = new DocManager(document);
        Map<String, String> map = new HashMap<>();
        map.put("NAME", "小明");
        map.put("AGE", "20");
        map.put("SEX", "男");
        Document target = docManager.fetchAField(map);

        File targetFile = new File("e:/test.docx");
        if (!targetFile.exists()) {
            targetFile.createNewFile();
        }
        OutputStream out = new FileOutputStream(targetFile);

        // 保存到流
        target.save(out, SaveFormat.DOCX);*/

        
      /*  try{
            Workbook wb = new Workbook("D:/test.xls");//打开模版文件
            WorkbookDesigner designer = new WorkbookDesigner();//加载设计器
            designer.setWorkbook(wb);
            designer.setDataSource("uname","李四");//设置变量数据
            designer.setDataSource("Score", getScore());//设置类对象数据
            designer.process();
            wb.save("E:/fex/test.xls");//生成报表
        }catch(Exception ex){
            ex.printStackTrace();
        }*/

        
    }

    @Test
    public void test1() throws Exception {
        InputStream input = Main.class.getClassLoader().getResourceAsStream("模板.docx");
        Document document = new Document(input);

        DocumentBuilder builder = new DocumentBuilder(document);

        builder.moveToBookmark("Student", true, true);
        
        //开始绘制表格
        Table table = builder.startTable();
        //插入一个单元格
        builder.insertCell();
        //设置表格缩进
        table.setLeftIndent(20.0);
        //设置行格式，宽度40，高度样式，至少40
        builder.getRowFormat().setHeight(40.0);
        builder.getRowFormat().setHeightRule(HeightRule.AT_LEAST);
        //设置背景颜色
        builder.getCellFormat().getShading().setBackgroundPatternColor(new Color(198, 217, 241));
        //段落垂直居中
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        //字体大小
        builder.getFont().setSize(16);
        //字体
        builder.getFont().setName("Arial");
        //加粗
        builder.getFont().setBold(true);
        //单元格宽度
        builder.getCellFormat().setWidth(100.0);
        //写入文字
        builder.write("Header Row,\n Cell 1");

        //第一行第二个单元格
        builder.insertCell();
        builder.write("Header Row,\n Cell 2");
        //第一行第三个单元格
        builder.insertCell();
        //设置宽度
        builder.getCellFormat().setWidth(200.0);
        builder.write("Header Row,\n Cell 3");
        //结束首行的绘制
        builder.endRow();

        //第二行
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.WHITE);
        builder.getCellFormat().setWidth(100.0);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        //行高度30，自动伸缩
        builder.getRowFormat().setHeight(30.0);
        builder.getRowFormat().setHeightRule(HeightRule.AUTO);
        //插入单元格
        builder.insertCell();
        builder.getFont().setSize(12);
        builder.getFont().setBold(false);
        //写入数据
        builder.write("Row 1, Cell 1 Content");
        //插入单元格
        builder.insertCell();
        //写入数据
        builder.write("Row 1, Cell 2 Content");
        //插入单元格
        builder.insertCell();
        //宽度
        builder.getCellFormat().setWidth(200.0);
        //写入数据
        builder.write("Row 1, Cell 3 Content");
        //第二行结束
        builder.endRow();

        //第三行
        builder.insertCell();
        builder.getCellFormat().setWidth(100.0);
        builder.write("Row 2, Cell 1 Content");

        builder.insertCell();
        builder.write("Row 2, Cell 2 Content");

        builder.insertCell();
        builder.getCellFormat().setWidth(200.0);
        builder.write("Row 2, Cell 3 Content.");
        builder.endRow();
        //结束表格的绘制
        builder.endTable();

        File targetFile = new File("e:/test2.docx");
        if (!targetFile.exists()) {
            targetFile.createNewFile();
        }
        OutputStream out = new FileOutputStream(targetFile);

        // 保存到流
        document.save(out, SaveFormat.DOCX);
    }

}
