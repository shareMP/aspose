package com.mrl.words;

import java.util.Map;
import java.util.Map.Entry;

import com.aspose.words.Document;

public class DocManager
{
    //一个文档对象
    private Document sourceDocument;

    public DocManager(Document sourceDocument) {
        this.sourceDocument = sourceDocument;
    }
    
    /**
     *  填充域
     *  [功能详细描述]
     *  @param list
     *  @return    
     * @exception/throws [违例类型] [违例说明]
     * @see [类、类#方法、类#成员]
     */
    public Document fetchAField(Map<String,String> map){
        
        if(map != null && map.size() > 0){
            
            //定义word中的域
            String[] fieldNames = new String[map.size()];
            //定义word域对应的值
            Object[] values = new Object[map.size()];
            
            //遍历map，把值放进去
            int num = 0;
            for(Entry<String,String> entry : map.entrySet()){
                fieldNames[num] = entry.getKey();
                values[num] = entry.getValue();
                num++;
            }
            
            //替换域对象
            try {
                sourceDocument.getMailMerge().execute(fieldNames, values);
            }
            catch (Exception e) {
                System.out.println("执行域代码的替换失败了！");
                e.printStackTrace();
            }
            
            return sourceDocument;
        }else{
            return null;
        }
    }
    
    
}
