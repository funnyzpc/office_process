package com.funny.office.service;

import com.funny.office.po.ProcessFileInfo;
import com.funny.office.utils.Excel2HtmlUtils;
import org.apache.commons.io.FileUtils;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.List;

@Service
public class ProcessExcelService {
    //@Autowired
    ScormService scormService=new ScormService();

    /**
     *
     * @param file       上传文件
     * @param uploadPath 目标文件存放文件夹
     * @return
     * @throws Exception
     */
    public ProcessFileInfo processXLSX(File file, String uploadPath)throws Exception {
        List<String> sheets= Excel2HtmlUtils.readExcelToHtml(file.getPath());
        FileUtils.forceMkdir(new File(uploadPath));//创建文件夹
        String code=file.getName().substring(0,file.getName().lastIndexOf("."));//文件名称
        String basePath=String.format("%s%s%s", uploadPath,File.separator,code);
        FileUtils.forceMkdir(new File(basePath));
        File htmlFile = new File(String.format("%s%s%s", basePath,File.separator,"index.html"));
        Writer fw=null;
        PrintWriter bw=null;
        //构建html文件
        try{
            fw= new BufferedWriter( new OutputStreamWriter(new FileOutputStream(htmlFile.getPath()),"UTF-8"));
            bw=new PrintWriter(fw);
            //添加表头及可缩放样式
            String head="<!DOCTYPE html><html><head><meta charset=\"UTF-8\"></head><body style=\"transform: scale(0.7,0.7);-webkit-transform: scale(0.7,0.7);\">";
            StringBuilder body=new StringBuilder();
            for (String e : sheets) {
                body.append(e);
            }
            String foot="</body></html>";
            bw.write(String.format("%s%s%s", head,body.toString(),foot));
        }catch(Exception e){
            throw new Exception("");//错误扔出
        }finally{
            if (bw != null) {
                bw.close();
            }
            if(fw!=null){
                fw.close();
            }
        }
        String htmlZipFile=String.format("%s%s%s.%s",uploadPath,File.separator,file.getName().substring(0,file.getName().lastIndexOf(".")),"ZIP");
        //压缩文件
        scormService.zip(basePath, htmlZipFile);
        file.delete();//删除上传的xlsx文件
        FileUtils.forceDelete(new File(basePath));
        return new ProcessFileInfo(true,new File(htmlZipFile).getName(),htmlZipFile);
    }

}
