package com.funny.office.service;

import com.funny.office.po.ProcessFileInfo;
import com.funny.office.service.ScormService;
import org.apache.commons.io.FileUtils;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class ProcessWordService {

    ScormService scormService=new ScormService();

    /**
     *
     * @param file       上传文件
     * @param uploadPath 目标文件存放文件夹
     * @return
     * @throws Exception
     */
    public ProcessFileInfo processDOCX(File file, String uploadPath) throws IOException {
        String fileName=file.getName().substring(0,file.getName().lastIndexOf("."));//获取文件名称
        String basePath=String.format("%s%s%s", uploadPath,File.separator,fileName);
        FileUtils.forceMkdir(new File(basePath));
        String zipFilePath=String.format("%s%s%s.%s", uploadPath,File.separator,fileName,"ZIP");
        FileOutputStream fos=null;
        try{
            fos=new FileOutputStream(new File(String.format("%s%s%s", basePath,File.separator,"index.html")));
            WordprocessingMLPackage wmp = WordprocessingMLPackage.load(file);//加载源文件
            Docx4J.toHTML(wmp, String.format("%s%s%s", basePath,File.separator,fileName),fileName,fos);
        }catch(Exception e){
            e.printStackTrace();
        }finally{
            fos.close();
        }
        scormService.zip(basePath, zipFilePath);
        // FileUtils.forceDelete(new File(String.format("%s%s%s", basePath,File.separator,"index.html")));
        FileUtils.forceDelete(new File(basePath));
        file.delete();
        return new ProcessFileInfo(true,new File(zipFilePath).getName(),zipFilePath);
    }
}
