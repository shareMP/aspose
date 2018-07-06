package com.mrl.words;

import java.util.Map;
import java.util.Map.Entry;

import com.aspose.words.Document;

public class DocManager
{
    //һ���ĵ�����
    private Document sourceDocument;

    public DocManager(Document sourceDocument) {
        this.sourceDocument = sourceDocument;
    }
    
    /**
     *  �����
     *  [������ϸ����]
     *  @param list
     *  @return    
     * @exception/throws [Υ������] [Υ��˵��]
     * @see [�ࡢ��#��������#��Ա]
     */
    public Document fetchAField(Map<String,String> map){
        
        if(map != null && map.size() > 0){
            
            //����word�е���
            String[] fieldNames = new String[map.size()];
            //����word���Ӧ��ֵ
            Object[] values = new Object[map.size()];
            
            //����map����ֵ�Ž�ȥ
            int num = 0;
            for(Entry<String,String> entry : map.entrySet()){
                fieldNames[num] = entry.getKey();
                values[num] = entry.getValue();
                num++;
            }
            
            //�滻�����
            try {
                sourceDocument.getMailMerge().execute(fieldNames, values);
            }
            catch (Exception e) {
                System.out.println("ִ���������滻ʧ���ˣ�");
                e.printStackTrace();
            }
            
            return sourceDocument;
        }else{
            return null;
        }
    }
    
    
}
