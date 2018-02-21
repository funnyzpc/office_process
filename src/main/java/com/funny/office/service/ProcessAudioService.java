package com.funny.office.service;

import com.funny.office.po.ProcessFileInfo;
import org.apache.commons.io.FileUtils;

import java.io.*;

public class ProcessAudioService {
    ScormService scormService=new ScormService();

    /**
     *
     * @param file						需要包装的文件
     * @param uploadPath		临时文件夹位置(upload的路径)
     * @param audioIMGPath	音频背景图片
     * @param audioIllustrated	音频简介
     * @return							返回处理状态(注意：包装成功后以当前code为包名)
     * @throws Exception
     */
    //写到这儿了：
    public ProcessFileInfo processAudio(File file, String uploadPath, String audioIMGPath, String audioIllustrated)throws Exception {
        /**
         * 功能逻辑
         * A>创建index文件并写入Html
         * B>构建ZIP文件将index.html以及音频文件写入ZIP
         * C>删除临时文件(index.html以及音频文件)
         * D>返回ZIP所在路径以及文件全名
         *
         */
        //A
        String code=file.getName().substring(0,file.getName().lastIndexOf("."));//文件名称
        String basePath=String.format("%s%s%s", uploadPath,File.separator,code);
        FileUtils.forceMkdir(new File(basePath));//创建base目录
        String suffix=file.getName().substring(file.getName().lastIndexOf(".")+1,file.getName().length()).toLowerCase();//音频文件后缀名
        File image=new File(audioIMGPath);
        StringBuilder html=new StringBuilder();
        html.append("<!DOCTYPE html><html><head><meta charset='utf-8'><title>Audio</title></head>");
        html.append("<body style=\"margin:0px 0px;\"><div style=\"width:100%;margin:auto 0% auto 0%;\">");
        html.append(String.format("%s%s%s","<video controls=\"controls\"  width=\"100%\"  height=\"100%\" name=\"media\" poster=\"",image.getName(),"\">"));//背景图片
        //多个类型,移动设备会取能播放的那个
        html.append(String.format("%s%s.%s%s","<source src=\"",code,suffix,"\"  type=\"audio/mp3\">"));
        html.append(String.format("%s%s.%s%s","<source src=\"",code,suffix,"\"  type=\"audio/wav\">"));
        html.append(String.format("%s%s.%s%s","<source src=\"",code,suffix,"\"  type=\"audio/m4a\">"));
        html.append(String.format("%s%s.%s%s","<source src=\"",code,suffix,"\"  type=\"audio/ogg\">"));
        html.append(String.format("%s%s.%s%s","<source src=\"",code,suffix,"\"  type=\"audio/acc\">"));
        html.append(String.format("%s%s.%s%s","<source src=\"",code,suffix,"\"  type=\"audio/mpeg\">"));
        html.append("</video></div>");
        if(null!=audioIllustrated){
            audioIllustrated=audioIllustrated.replaceAll("\r","<br/>");
            audioIllustrated=audioIllustrated.replaceAll("  ","&nbsp;&nbsp;");
            html.append(String.format("%s%s%s", "<div style=\"margin:5% 5% 5% 5%;font-size:18px;padding-bottom:10%;\">",audioIllustrated,"</div>"));//简介
        }
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
        //B
        String zipPath = String.format("%s%s%s.%s", uploadPath,File.separator,code, "ZIP");
        //音频文件拷贝到basePath中
        FileUtils.moveFile(file, new File(String.format("%s%s%s",basePath,File.separator,file.getName())));
        if("audio.png".equals(image.getName())){
            File imageFile=new File(String.format("%s%s%s", basePath,File.separator, "audio.png"));
            FileUtils.copyFile(image, imageFile);
            //Files.copy(new File(audioIMGPath).toPath(), imageFile.toPath());//将默认图片文件复制到upload/basePath内
        }else{
            FileUtils.moveFile(image, new File(String.format("%s%s%s",basePath,File.separator,image.getName())));
            //Files.copy(new File(audioIMGPath).toPath(), new File(String.format("%s%s%s", basePath,File.separator,image.getName())).toPath());
        }
        scormService.zip(basePath, zipPath);//打包文件
        //file.delete();//删除音频文件
        FileUtils.forceDelete(new File(basePath));
        return new ProcessFileInfo(true,String.format("%s.%s",code,"ZIP"),zipPath);
    }
}
