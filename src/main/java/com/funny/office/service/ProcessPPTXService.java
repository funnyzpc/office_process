package com.funny.office.service;

import com.funny.office.po.ProcessFileInfo;
import org.apache.commons.io.FileUtils;
import org.springframework.stereotype.Service;

import java.io.*;

@Service
public class ProcessPPTXService {
    ScormService scormService=new ScormService();

    public ProcessFileInfo processPPTX(File file, String uploadPath)throws Exception{
        String fileName=file.getName().substring(0,file.getName().lastIndexOf("."));//获取文件名称
        String suffix=file.getName().substring(file.getName().lastIndexOf(".")+1,file.getName().length()).toLowerCase();//音频文件后缀名
        String basePath=String.format("%s%s%s", uploadPath,File.separator,fileName);
        FileUtils.forceMkdir(new File(basePath));
        //将视频文件copy到basePath内
        String videoPath=String.format("%s%s%s", basePath,File.separator,file.getName());
        FileUtils.copyFile(file, new File(videoPath));
        StringBuilder html=new StringBuilder();
        html.append("<!DOCTYPE html><html><head><meta charset='utf-8'><title>powerpoint</title></head>");
        html.append("<body style=\"margin:0px 0px;\"><div style=\"width:100%;margin:auto 0% auto 0%;\">");
        html.append("<video controls=\"controls\"  width=\"100%\"  height=\"100%\" name=\"media\" >");//无背景图片
        html.append(String.format("%s%s.%s%s%s%s%s","<source src=\"",fileName,suffix,"\" type=\"audio/",suffix,"\" >","</video></div>"));//视频
        html.append("</body></html>");//结尾
        File indexFile=new File(String.format("%s%s%s",basePath,File.separator,"index.html"));
        Writer fw=null;
        PrintWriter bw=null;
        //构建文件(html写入html文件)
        try{
            fw= new BufferedWriter( new OutputStreamWriter(new FileOutputStream(indexFile),"UTF-8"));//以UTF-8的格式写入文件
            bw=new PrintWriter(fw);
            bw.write(html.toString());
        }catch(Exception e){
            throw new Exception(e.toString());//错误扔出
        }finally{
            if (bw != null) {
                bw.close();
            }
            if(fw!=null){
                fw.close();
            }
        }
        String zipFilePath=String.format("%s%s%s.%s", uploadPath,File.separator,fileName,"ZIP");
        scormService.zip(basePath, zipFilePath);
        //删除文件
        file.delete();
        FileUtils.forceDelete(new File(basePath));
        return new ProcessFileInfo(true,new File(zipFilePath).getName(),zipFilePath);
    }
}
