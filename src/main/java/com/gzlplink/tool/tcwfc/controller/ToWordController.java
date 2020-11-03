package com.gzlplink.tool.tcwfc.controller;

import cn.afterturn.easypoi.word.WordExportUtil;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 生成word
 *
 * @author wangbinbin
 * @date 2020-09-04 14:24:30
 */
@Slf4j
@RestController
@Api(tags = "生成word")
public class ToWordController {
    private static List<String> filePaths=new ArrayList<>();
    private static Integer[] fileLines=null;
    private static ConcurrentHashMap<Integer,String> fileContext=new ConcurrentHashMap<Integer,String>();


    @ApiOperation(value = "生成源代码word")
    @RequestMapping(value = "/toword", method = RequestMethod.POST)
    public void delete(
            @RequestParam(value = "sourceDir") @ApiParam(value = "源代码目录", required = true,example = "D:\\dev\\workspace\\lepeng\\mps-edp\\mps-edp-srv\\src")  String sourceDir
            ,@RequestParam(value = "extNames") @ApiParam(value = "源代码文件扩展名，英文逗号分割，如.java,.xml", required = true)  List<String> extNames
            ,@RequestParam(value = "sysName") @ApiParam(value = "系统名称", required = true,example = "电算化报表",defaultValue = "系统名称")  String sysName
            ,@RequestParam(value = "sysDesc") @ApiParam(value = "功能描述", required = true,example = "提供支撑公司各类业务的财务相关的各类报表",defaultValue = "功能描述")  String sysDesc
            ,@RequestParam(value = "codeLineNum") @ApiParam(value = "生成代码行数", required = true,example = "3000",defaultValue = "3000")  Integer codeLineNum
            ,@RequestParam(value = "wordPath") @ApiParam(value = "生成word全路径", required = true,example = "D:\\电算化报表.docx",defaultValue = "D:\\源代码文档.docx")  String wordPath
    ) {
        toWord(sourceDir,sysName,sysDesc,"templates\\code2word.docx",wordPath,codeLineNum,extNames);
    }

    /**
     * 生成WORD <br>
     * @author wangbinbin
     */
    private static void toWord(String sourceDir,String sysName,String sysDesc,String templatePath,String wordPath,int codeLineNum,List<String> extNames) {
        filePaths.clear();
        fileContext.clear();

        readfile(sourceDir);
        final int HALF_LINE_NUM=codeLineNum/2;
        if(filePaths!=null && filePaths.size()>0) {
            fileLines=new Integer[filePaths.size()];
            int fileNo=0;
            StringBuilder sb=null;
            for(String filePath:filePaths){
                sb=new StringBuilder();
                int lineNo = 0;
                try {
                    String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);
                    String extName = filePath.substring(filePath.lastIndexOf(".")).toLowerCase();
                    if(filePath.toLowerCase().indexOf("test")==-1 && extNames.contains(extName)){
                        sb.append("\r//========== 文件").append(fileName).append("的源代码 ==========\r");
                        BufferedReader in = new BufferedReader(new FileReader(filePath));
                        String str;
                        lineNo = 2;
                        while ((str = in.readLine()) != null) {
                            if (str != null && !"".equals(str.trim())) {
                                sb.append(str).append("\r");
                                lineNo++;
                            }
                        }
                    }
                } catch (Exception e) {
                }
                fileLines[fileNo]=lineNo;
                fileContext.put(fileNo,sb.toString());
                fileNo++;
            }
        }

        int lineSum=0;
        int middleFileNo=0;
        for(int i=fileLines.length-1;lineSum<HALF_LINE_NUM && i>=0;i--){
            lineSum+=fileLines[i].intValue();
            middleFileNo=i--;
        }

        StringBuilder text=new StringBuilder();
        lineSum=0;

        //首部1500行
        for(int i=0;lineSum<HALF_LINE_NUM;i++){
            lineSum+=fileLines[i].intValue();
            text.append(fileContext.get(i));
        }

        //尾部1500行
        for(int i=middleFileNo;i<fileLines.length;i++){
            text.append(fileContext.get(i));
        }
        Map<String, Object> map = new HashMap<>();
        map.put("sysName", sysName);
        map.put("sysDesc", sysDesc);
        map.put("sourceCode", text.toString());
        //----------------------------------------------
        try {
            XWPFDocument doc = WordExportUtil.exportWord07(templatePath, map);
            FileOutputStream fos = new FileOutputStream(wordPath);
            doc.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * 读取某个文件夹下的所有文件(支持多级文件夹)
     */
    public static boolean readfile(String filepath){
        try {
            File file = new File(filepath);
            if (!file.isDirectory()) {
                filePaths.add(file.getAbsolutePath());
            } else if (file.isDirectory()) {
                String[] filelist = file.list();
                for (int i = 0; i < filelist.length; i++) {
                    File readfile = new File(filepath + "\\" + filelist[i]);
                    if (!readfile.isDirectory()) {
                        filePaths.add(readfile.getAbsolutePath());
                    } else if (readfile.isDirectory()) {
                        readfile(filepath + "\\" + filelist[i]);
                    }
                }
            }
        } catch (Exception e) {
            log.error("readfile()   Exception:" + e.getMessage());
        }
        return true;
    }

}
