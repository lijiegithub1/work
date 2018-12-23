package cn.jd.web.controller;


import org.apache.commons.lang.WordUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Administrator on 2018/07/07.
 * 初始化的controller
 */
@Controller

public class IndexController {
    @Autowired
    private SummaryExpService summaryExpServiceImpl;

    @RequestMapping("index")
    public String showIndex(HttpServletRequest request){
        return "index";
    }

    @RequestMapping("importExp")
    @ResponseBody
    public String importExp(@RequestParam("file") MultipartFile file, HttpServletRequest request){
        // 判断文件是否为空
        String flag = "02";//上传标志
        if (!file.isEmpty()) {
            try {
                String originalFilename = file.getOriginalFilename();//原文件名字
                InputStream is = file.getInputStream();//获取输入流
               // Map<String,Object> datas= summaryExpServiceImpl.writeExelData(is);
               Map<String,Object> datas = new HashMap<String, Object>();
               // datas.put("name","羽绒服");
                WordDocumentService wordDocumentService = new WordDocumentService();
                wordDocumentService.datasMappingWord(datas,"人事变动表 (2).ftl","H://word//自动生成的文档.doc");
                // 转存文件
                //file.transferTo(new File(filePath));
            } catch (Exception e) {
                flag="03";//上传出错
                e.printStackTrace();
            }
        }
        return flag;
    }

    public static void main(String[] args) {
        SummaryExpService summaryExpServiceImpl = new SummaryExpServiceImpl();
        String filePath = "H:\\小碗\\excel\\职位变动表-2018 (1).xlsx";
        Map<String,Object> datas = null;
        try {
              summaryExpServiceImpl.writeExelData(new FileInputStream(filePath));
        }catch (Exception e){
            System.out.println(e);
        }
    }
}
