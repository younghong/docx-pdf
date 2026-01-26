package com.young;

import java.io.File;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;

public class App {

    public static void main(String[] args) throws Exception {
    	//convertvx();
    	convertv1();
    	//convertv2();
    	//test 
    }
    
    public static void convertvx()
    {
    	System.out.println("test branch");
    }
    
    public static void convertv1()
    {
    	long startTime = System.currentTimeMillis(); // 코드 시작 시간

    	
    	
		String mainfilepath = "D:\\dev\\003_study\\docxsample\\test001\\";
		
//		String filename="한국가스공사";
//		String filename="상세침투결과";
		String filename="error";
		
        String inputfilepath = mainfilepath + filename+ ".docx";
        String outputfilepath = mainfilepath + "output912_"+filename+".pdf";

        String chageoutputfilepath = mainfilepath + "changeoutput_"+filename+".pdf";
        
        String fopConfigPath = "D:\\dev\\003_study\\docxsample\\fop.xconf";
    	
    	
    	docx2pdf con = new docx2pdf();
    	
    	con.toPDF(inputfilepath, outputfilepath, fopConfigPath);
    	
    	
    	
    	
    	
//    	PDDocument doc;
//		try {
//			doc = PDDocument.load(new File(outputfilepath));
//			
//	    	PDDocumentInformation info = doc.getDocumentInformation();
//
//	    	info.setProducer("MyCompany PDF Engine");
//	    	info.setCreator("My Converter");
//
//	    	doc.save(chageoutputfilepath);
//	    	doc.close();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}


    	
    	
    	
    	
    	
    	long endTime = System.currentTimeMillis(); // 코드 끝난 시간
    	long durationTimeSec = endTime - startTime;
        
    	System.out.println(durationTimeSec + "m/s"); // 밀리세컨드 출력
    	System.out.println((durationTimeSec / 1000) + "sec"); // 초 단위 변환 출력

    }
    
    
}