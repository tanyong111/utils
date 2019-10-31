package com.tan.utils.rap2toword;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PaseRap2JsonUtils {

    public static Map<String,Object> readJson(String path) {
        File file=new File(path);
        if(!file.exists()) {
            return null;
        }
        try(FileReader reader = new FileReader(file);
            BufferedReader bReader = new BufferedReader(reader);) {
            StringBuilder sb = new StringBuilder();//定义一个字符串缓存，将字符串存放缓存中
            String s = "";
            while ((s =bReader.readLine()) != null) {//逐行读取文件内容，不读取换行符和末尾的空格
                sb.append(s + "\n");//将读取的字符串添加换行符后累加存放在缓存中
                System.out.println(s);
            }
            String str = sb.toString();
            System.out.println(str);
            ObjectMapper objectMapper = new ObjectMapper();
            return objectMapper.readValue(str, new TypeReference<Map<String,Object>>() {
            });
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public static List<Map<String,Object>> toWordData(Map<String,Object> data) {
        List<Map<String,Object>> wordData=new ArrayList<>();
        List<Map<String,Object>> modules=(List<Map<String,Object>>)((Map<String,Object>)data.get("data")).get("modules");
        for(Map<String,Object> module:modules) {
            Map<String,Object> wordModule=new HashMap<>();
            String moduleName=(String) module.get("name");
            String moduleDescription=(String) module.get("description");
            wordModule.put("name",moduleName);
            wordModule.put("description",moduleDescription);
            List<Map<String,Object>> interfaces=(List<Map<String,Object>>) module.get("interfaces");
            List<Map<String,Object>> wordInterfaces=new ArrayList<>();
            for(Map<String,Object> interf:interfaces) {
                Map<String,Object> wordInterf=new HashMap<>();
                String interfName=(String) interf.get("name");
                String interfUrl=(String) interf.get("url");
                String interfMethod=(String) interf.get("method");
                String interfDescription=(String) interf.get("description");
                Integer interfStatus=(Integer) interf.get("status");
                wordInterf.put("name",interfName);
                wordInterf.put("url",interfUrl);
                wordInterf.put("method",interfMethod);
                wordInterf.put("description",interfDescription);
                wordInterf.put("status",interfStatus);
                List<Map<String,Object>> interfProperties=(List<Map<String,Object>>) interf.get("properties");
                List<Map<String,Object>> request = new ArrayList<>();
                List<Map<String,Object>> response = new ArrayList<>();
                for(Map<String,Object> interfPropertie:interfProperties) {
                    Map<String,Object> propertis=new HashMap<>();
                    propertis.put("id",interfPropertie.get("id"));
                    propertis.put("scope",interfPropertie.get("scope"));
                    propertis.put("type",interfPropertie.get("type"));
                    propertis.put("name",interfPropertie.get("name"));
                    propertis.put("rule",interfPropertie.get("rule"));
                    propertis.put("value",interfPropertie.get("value"));
                    propertis.put("description",interfPropertie.get("description"));
                    propertis.put("parentId",interfPropertie.get("parentId"));
                    propertis.put("required",interfPropertie.get("required"));
                    if("request".equals(propertis.get("scope"))) {
                        if((Integer)propertis.get("parentId") == -1) {
                            request.add(propertis);
                        } else {
                            // 递归遍历所有树节点，直到找到其父节点
                            appendToParent(request,propertis);
                        }
                    } else if("response".equals(propertis.get("scope"))) {
                        if((Integer)propertis.get("parentId") == -1) {
                            response.add(propertis);
                        } else {
                            // 递归遍历所有树节点，直到找到其父节点
                            appendToParent(response,propertis);
                        }
                    }
                }
                wordInterf.put("request",request);
                wordInterf.put("response",response);
                wordInterfaces.add(wordInterf);
            }
            wordModule.put("interfaces",wordInterfaces);
            wordData.add(wordModule);
        }
        return wordData;
    }

    private static void appendToParent(List<Map<String,Object>> parents,Map<String,Object> propertis) {
        for(Map<String,Object> parent:parents) {
            if(parent.get("id").equals(propertis.get("parentId"))) {
                if(parent.get("children") == null) {
                    parent.put("children",new ArrayList<Map<String,Object>>());
                }
                List<Map<String,Object>> children=(List<Map<String,Object>>)parent.get("children");
                children.add(propertis);
                return;
            } else if(parent.get("children")!=null) {
                appendToParent((List<Map<String,Object>>)parent.get("children"),propertis);
            }
        }
    }

    // 设置模块标题
    private static void setModuleTitleParagraph(XWPFDocument document,String name) {
        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText(name);
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);
        titleParagraphRun.setBold(true);
        addCustomHeadingStyle(document,"模块标题",1);
        titleParagraph.setStyle("模块标题");
    }

    // 设置接口标题
    private static void setInterfacesTitleParagraph(XWPFDocument document,String name) {
        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText(name);
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(14);
        titleParagraphRun.setBold(true);
        addCustomHeadingStyle(document,"接口标题",2);
        titleParagraph.setStyle("接口标题");
    }

    // 设置普通段落
    private static void setParagraph(XWPFDocument document,String text) {
        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText(text);
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(12);
    }

    /**
     * 增加自定义标题样式。这里用的是stackoverflow的源码
     *
     * @param docxDocument 目标文档
     * @param strStyleId 样式名称
     * @param headingLevel 样式级别
     */
    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);

    }

    private static void setInterfProperties(XWPFDocument document,List<Map<String,Object>> request) {
        //换行
        //XWPFParagraph paragraph1 = document.createParagraph();
        //XWPFRun paragraphRun1 = paragraph1.createRun();
        //paragraphRun1.setText("\r");
        //基本信息表格
        XWPFTable infoTable = document.createTable();
        //去表格边框
        infoTable.getCTTbl().getTblPr().unsetTblBorders();
        //列宽自动分割
        CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
        infoTableWidth.setType(STTblWidth.DXA);
        infoTableWidth.setW(BigInteger.valueOf(9072));

        //表格第一行
        XWPFTableRow infoTableRowOne = infoTable.getRow(0);
        infoTableRowOne.getCell(0).setText("名称");
        infoTableRowOne.addNewTableCell().setText("必选");
        infoTableRowOne.addNewTableCell().setText("类型");
        infoTableRowOne.addNewTableCell().setText("生成规则");
        infoTableRowOne.addNewTableCell().setText("初始值");
        infoTableRowOne.addNewTableCell().setText("简介");
        // 深度遍历添加参数
        setTable(infoTable,request,"");
    }

    private static void setTable(XWPFTable infoTable,List<Map<String,Object>> parents,String prifix) {
        for(Map<String,Object> parent:parents) {
            XWPFTableRow infoTableRowTwo = infoTable.createRow();
            infoTableRowTwo.getCell(0).setText(prifix+ toString(parent.get("name")));
            infoTableRowTwo.getCell(1).setText(prifix+ toString(parent.get("required")));
            infoTableRowTwo.getCell(2).setText(prifix+toString(parent.get("type")));
            infoTableRowTwo.getCell(3).setText(prifix+toString(parent.get("rule")));
            infoTableRowTwo.getCell(4).setText(prifix+toString(parent.get("value")));
            infoTableRowTwo.getCell(5).setText(prifix+toString(parent.get("description")));
            List<Map<String,Object>> childrens=(List<Map<String,Object>>)parent.get("children");
            String pri=prifix+" ";
            if(childrens!=null) {
                setTable(infoTable,childrens,pri);
            }
        }
    }

    private static String toString(Object str) {
        if(str!=null) {
            return str.toString();
        } else {
            return "";
        }
    }

    /**
     * 1. 某某接口，
     * 请求地址：
     * 请求方法：
     * 请求参数：
     * 请求结果：（json）
     * @param wordData
     */
    public static void toWord(List<Map<String,Object>> wordData) {
        System.out.println(wordData);
        XWPFDocument document = new XWPFDocument();
        for(int i=0;i<wordData.size();i++) {
            Map<String,Object> module=wordData.get(i);
            // 按模块分章节
            String moduleName = (String) module.get("name");
            //添加标题
            setModuleTitleParagraph(document,(i+1)+"."+moduleName);
            // 添加接口
            List<Map<String,Object>> interfaces=(List<Map<String,Object>>) module.get("interfaces");
            for(int j=0;j<interfaces.size();j++) {
                Map<String,Object> interf=interfaces.get(j);
                String interfName=(String) interf.get("name");
                String interfUrl=(String) interf.get("url");
                String interfMethod=(String) interf.get("method");
                String interfDescription=(String) interf.get("description");
                Integer interfStatus=(Integer) interf.get("status");
                List<Map<String,Object>> request = (List<Map<String,Object>>)interf.get("request");
                List<Map<String,Object>> response = (List<Map<String,Object>>)interf.get("response");
                setInterfacesTitleParagraph(document,(i+1)+"."+(j+1)+" "+interfName);
                setParagraph(document,"请求地址："+interfUrl);
                setParagraph(document,"请求方法："+interfMethod);
                setParagraph(document,"简介："+interfDescription);
                setParagraph(document,"请求参数：");
                setInterfProperties(document,request);
                setParagraph(document,"请求结果：");
                setInterfProperties(document,response);
            }
        }

        // 输出到word
        File file = new File("D:\\t.docx");
        FileOutputStream out= null;
        try {
            out = new FileOutputStream(file);
            document.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) throws IOException {
        Map<String,Object> data=readJson("D:\\工作\\零度视界\\招行银团项目\\项目交接文件\\rap2.json");
        List<Map<String,Object>> wordData=toWordData(data);
        toWord(wordData);
    }
}
