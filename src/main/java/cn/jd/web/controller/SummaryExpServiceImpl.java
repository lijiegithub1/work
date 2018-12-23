package cn.jd.web.controller;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class SummaryExpServiceImpl implements SummaryExpService {

    @Override
    public void writeExelData(InputStream is) {
        XSSFWorkbook xssfWorkbook = null;
        Map<String,Object> datas = new HashMap<String, Object>();
        try {
            xssfWorkbook = new XSSFWorkbook(is);
        } catch (IOException e) {
            System.out.println("e-->" + e);
        }
        if (xssfWorkbook == null) {
            System.out.println("未读取到内容,请检查路径！");

        }
        //遍历xlsx中的sheet
        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
            String sheetName = xssfWorkbook.getSheetName(numSheet);
            if(!StringUtils.equals("2018.12",sheetName)){
                continue;
            }

            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
            if (xssfSheet == null) {
                continue;
            }
            // 对于每个sheet，读取其中的每一行
            for (int rowNum = 2; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                if (xssfRow == null) continue;
                String name = Trim_str( getValue(xssfRow.getCell(7)));
                datas.put("name",name );
                datas.put("depart2",Trim_str( getValue(xssfRow.getCell(1))));
                datas.put("depart1",Trim_str( getValue(xssfRow.getCell(3))));
                datas.put("position",Trim_str( getValue(xssfRow.getCell(8))));
                datas.put("date1",Trim_str( getValue(xssfRow.getCell(14))));
                datas.put("num",Trim_str( getValue(xssfRow.getCell(6))));
                datas.put("effDate",Trim_str( getValue(xssfRow.getCell(16))));
                datas.put("trialSalary",Trim_str( getValue(xssfRow.getCell(10))));
                datas.put("regularSalary",Trim_str( getValue(xssfRow.getCell(13))));
                WordDocumentService wordDocumentService = new WordDocumentService();
                wordDocumentService.datasMappingWord(datas,"人事变动表.ftl","H:\\小碗\\生成word\\"+name+".doc");
            }
        }
    }

    //判断后缀为xlsx的excel文件的数据类
    @SuppressWarnings("deprecation")
    private static String getValue(XSSFCell xssfRow) {
        if(xssfRow==null){
            return "---";
        }
        if (xssfRow.getCellType() == xssfRow.CELL_TYPE_BOOLEAN) {
            return String.valueOf(xssfRow.getBooleanCellValue());
        } else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_NUMERIC) {
            String strCell = "";
            if (XSSFDateUtil.isCellDateFormatted(xssfRow)) {
                //  如果是date类型则 ，获取该cell的date值
                strCell = new SimpleDateFormat("yyyy/MM/dd").format(XSSFDateUtil.getJavaDate(xssfRow.getNumericCellValue()));
            } else { // 纯数字 转成Integer
                strCell = String.valueOf(Double.valueOf(xssfRow.getNumericCellValue()).intValue());
            }

            return strCell;
        } else if(xssfRow.getCellType() == xssfRow.CELL_TYPE_BLANK || xssfRow.getCellType() == xssfRow.CELL_TYPE_ERROR){
            return "---";
        }
        else {//字符串
            return String.valueOf(xssfRow.getStringCellValue());
        }
    }

    //字符串修剪  去除所有空白符号 ， 问号 ， 中文空格
    static private String Trim_str(String str){
        if(str==null)
            return null;
        return str.replaceAll("[\\s\\?]", "").replace("　", "");
    }

}
































