package cn.jd.web.controller;

import freemarker.template.Configuration;
import freemarker.template.Template;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Map;

/**
 * Word文档业务服务
 * Created by louie on 2017-11-13.
 */
public class WordDocumentService {
    private Configuration configuration ;

    public WordDocumentService(){
        this.configuration = new Configuration(Configuration.VERSION_2_3_23);
        configuration.setDefaultEncoding("utf8");
      //  configuration.setClassForTemplateLoading(this.getClass(),"/");
        try {
            configuration.setDirectoryForTemplateLoading(new File("H:\\小碗\\word"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void datasMappingWord(Map<String,Object> datas,String templateName,String outWordName){
        try {
            Template template = configuration.getTemplate(templateName,"utf8");
            File wordFile = new File(outWordName);
            FileWriter writer = new FileWriter(wordFile);
            template.process(datas,writer);

            writer.flush();
            writer.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
