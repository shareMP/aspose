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

        // ע����
        // String licenseName = Main.class.getResource("/").getPath() +
        // "license.xml";
        // String licenseName =
        // "D:/���˹����ռ�/study/words/src/main/resources/license.xml";
       /* InputStream inputLincese = Main.class.getClassLoader().getResourceAsStream("license.xml");
        License license = new License();
        license.setLicense(inputLincese);

        InputStream input = Main.class.getClassLoader().getResourceAsStream("ģ��.docx");
        Document document = new Document(input);
        DocManager docManager = new DocManager(document);
        Map<String, String> map = new HashMap<>();
        map.put("NAME", "С��");
        map.put("AGE", "20");
        map.put("SEX", "��");
        Document target = docManager.fetchAField(map);

        File targetFile = new File("e:/test.docx");
        if (!targetFile.exists()) {
            targetFile.createNewFile();
        }
        OutputStream out = new FileOutputStream(targetFile);

        // ���浽��
        target.save(out, SaveFormat.DOCX);*/

        
      /*  try{
            Workbook wb = new Workbook("D:/test.xls");//��ģ���ļ�
            WorkbookDesigner designer = new WorkbookDesigner();//���������
            designer.setWorkbook(wb);
            designer.setDataSource("uname","����");//���ñ�������
            designer.setDataSource("Score", getScore());//�������������
            designer.process();
            wb.save("E:/fex/test.xls");//���ɱ���
        }catch(Exception ex){
            ex.printStackTrace();
        }*/

        
    }

    @Test
    public void test1() throws Exception {
        InputStream input = Main.class.getClassLoader().getResourceAsStream("ģ��.docx");
        Document document = new Document(input);

        DocumentBuilder builder = new DocumentBuilder(document);

        builder.moveToBookmark("Student", true, true);
        
        //��ʼ���Ʊ��
        Table table = builder.startTable();
        //����һ����Ԫ��
        builder.insertCell();
        //���ñ������
        table.setLeftIndent(20.0);
        //�����и�ʽ�����40���߶���ʽ������40
        builder.getRowFormat().setHeight(40.0);
        builder.getRowFormat().setHeightRule(HeightRule.AT_LEAST);
        //���ñ�����ɫ
        builder.getCellFormat().getShading().setBackgroundPatternColor(new Color(198, 217, 241));
        //���䴹ֱ����
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        //�����С
        builder.getFont().setSize(16);
        //����
        builder.getFont().setName("Arial");
        //�Ӵ�
        builder.getFont().setBold(true);
        //��Ԫ����
        builder.getCellFormat().setWidth(100.0);
        //д������
        builder.write("Header Row,\n Cell 1");

        //��һ�еڶ�����Ԫ��
        builder.insertCell();
        builder.write("Header Row,\n Cell 2");
        //��һ�е�������Ԫ��
        builder.insertCell();
        //���ÿ��
        builder.getCellFormat().setWidth(200.0);
        builder.write("Header Row,\n Cell 3");
        //�������еĻ���
        builder.endRow();

        //�ڶ���
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.WHITE);
        builder.getCellFormat().setWidth(100.0);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        //�и߶�30���Զ�����
        builder.getRowFormat().setHeight(30.0);
        builder.getRowFormat().setHeightRule(HeightRule.AUTO);
        //���뵥Ԫ��
        builder.insertCell();
        builder.getFont().setSize(12);
        builder.getFont().setBold(false);
        //д������
        builder.write("Row 1, Cell 1 Content");
        //���뵥Ԫ��
        builder.insertCell();
        //д������
        builder.write("Row 1, Cell 2 Content");
        //���뵥Ԫ��
        builder.insertCell();
        //���
        builder.getCellFormat().setWidth(200.0);
        //д������
        builder.write("Row 1, Cell 3 Content");
        //�ڶ��н���
        builder.endRow();

        //������
        builder.insertCell();
        builder.getCellFormat().setWidth(100.0);
        builder.write("Row 2, Cell 1 Content");

        builder.insertCell();
        builder.write("Row 2, Cell 2 Content");

        builder.insertCell();
        builder.getCellFormat().setWidth(200.0);
        builder.write("Row 2, Cell 3 Content.");
        builder.endRow();
        //�������Ļ���
        builder.endTable();

        File targetFile = new File("e:/test2.docx");
        if (!targetFile.exists()) {
            targetFile.createNewFile();
        }
        OutputStream out = new FileOutputStream(targetFile);

        // ���浽��
        document.save(out, SaveFormat.DOCX);
    }

}
