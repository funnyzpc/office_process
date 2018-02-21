package com.funny.office.rest;

import com.funny.office.service.*;

import java.io.File;

public class Test {
    static ProcessExcelService processExcelServie;
    static ProcessWordService processWordService;
    static ProcessPDFService processPDFService;
    static ProcessPPTXService processPPTXService;
    static ProcessAudioService processAudioService;

    public static void main(String[] args)throws Exception {
        /**此处为具体调用方式(未测试)
         * 如有疑问可通过以下三种方式解决=>
         *  A>http://www.cnblogs.com/funnyzpc/p/7225988.html
         *
         *  B>调试+Google
         *
         *  C>电邮咨询:funnyzpc@gmail.com
         */
        /**
                                      XLSX	->	XLSX
         * 								ZIP		->	PDF(图片包)
         * 								DOCX	->	DOCX
         * 								MP4		->	PPTX(幻灯片转MP4)
         */
        //Excel解析
        processExcelServie=new ProcessExcelService();
        processExcelServie.processXLSX(new File("/var/cc.xlsx"),"/var/output");

        //Word解析
        processWordService=new ProcessWordService();
        processWordService.processDOCX(new File("/var/cc.docx"),"/var/output");

        //PowerPoint解析
        processPPTXService=new ProcessPPTXService();
        processPPTXService.processPPTX(new File("/var/cc.mp4"),"/var/output");

        //PDF解析
        processPDFService=new ProcessPDFService();
        processPDFService.processPDF(new File("/var/cc.zip"),"/var/output");

        //audio类文件包装(具体调用方式请见processAudio参数说明)
        processAudioService=new ProcessAudioService();
        //processAudioService.processAudio();
    }


}
