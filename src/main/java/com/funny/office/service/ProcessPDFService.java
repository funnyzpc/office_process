package com.funny.office.service;

import com.funny.office.po.ProcessFileInfo;
import org.apache.commons.io.FileUtils;

import java.io.*;
import java.util.*;
import java.util.regex.Pattern;

public class ProcessPDFService {
    ScormService scormService=new ScormService();

    public ProcessFileInfo processPDF(File file, String uploadPath)throws Exception{
        String fileName=file.getName().substring(0,file.getName().lastIndexOf("."));//获取文件名称
        String basePath=String.format("%s%s%s",uploadPath,File.separator,fileName);
        String unZipPath=String.format("%s%s%s%s",basePath,File.separator,fileName,File.separator);
        FileUtils.forceMkdir(new File(unZipPath));
        scormService.unzip(file.getPath(), unZipPath);//加压Zip文件
        //遍历文件夹获取文件名
        File[] f=new File(unZipPath).listFiles();
        List<String> imgNames=new ArrayList<String>();
        for (File ff : f) {
            imgNames.add(ff.getName());
        }

        /**
         * 根据文件名中的数字排列图片
         * 	a>提取文件名中的数字放入int数组(序列)
         *  b>判断序列数组元素个数与文件个数是否一致,不一致则抛出
         *  c>将序列数组从小到大排列
         *  d>遍历序列数组获取Map中的文件名(value)并写html
         */
        String nm=null;
        int[] i=new int[imgNames.size()];
        Map<Integer,String> names=new HashMap<Integer,String>();
        Pattern p=Pattern.compile("[^0-9]");
        for(int j=0;j<imgNames.size();j++){
            nm=imgNames.get(j).substring(0,imgNames.get(j).lastIndexOf("."));//提取名称
            String idx=p.matcher(nm).replaceAll("").trim();
            i[j]=Integer.parseInt("".equals(idx)?"0":idx);
            names.put(i[j],imgNames.get(j));
        }
        if(names.keySet().size()!=i.length){
            //System.out.println("====请检查您的图片编号====");/*重复或者不存在数字编号*/
            return new ProcessFileInfo(false,null,null);
        }
        Arrays.sort(i);//int数组内元素从小到大排列

        //包装成html
        StringBuilder html=new StringBuilder();
        html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'><title>PDF</title></head>");
        html.append("<body style=\"margin:0px 0px;padding:0px 0px;\">");
        for (int  k : i) {
            html.append(String.format("%s%s%s%s%s","<div style=\"width:100%;\"><img src=\"./",fileName,File.separator,names.get(k),"\"  style=\"width:100%;\" /></div>"));
        }
        html.append("</body></html>");
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
        String zipFilePath=String.format("%s%s%s.%s", uploadPath,File.separator,file.hashCode(),"ZIP");
        scormService.zip(basePath, zipFilePath);
        //删除文件
        file.delete();
        FileUtils.forceDelete(new File(basePath));
        return new ProcessFileInfo(true,new File(zipFilePath).getName(),zipFilePath);
    }
}
