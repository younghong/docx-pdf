package com.young;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.pdfbox.multipdf.PDFMergerUtility;

public class TotalApp {

	public static void main(String[] args) {
		
		long startTime = System.currentTimeMillis(); // 코드 시작 시간
		
		
		/* file 경로 설정. */
		String mainfilepath = "D:\\dev\\003_study\\docxsample\\test001\\";
		String filename="한국가스공사";
        String inputfilepath = mainfilepath + filename+ ".docx";
        String fopConfigPath = "D:\\dev\\003_study\\docxsample\\fop.xconf";
		
		
		docxSplit2pdf d2p=new docxSplit2pdf();
		d2p.toPDF(inputfilepath, mainfilepath, fopConfigPath);
		
		
		
    	long endTime = System.currentTimeMillis(); // 코드 끝난 시간
    	long durationTimeSec = endTime - startTime;
        
    	System.out.println(durationTimeSec + "m/s"); // 밀리세컨드 출력
    	System.out.println((durationTimeSec / 1000) + "sec"); // 초 단위 변환 출력
	}
	
	
	
	public static void convet()
	{
		/* file 경로 설정. */
		String mainfilepath = "D:\\dev\\003_study\\docxsample\\test001\\";
		String filename="한국가스공사";
        String inputfilepath = mainfilepath + filename+ ".docx";
        String fopConfigPath = "D:\\dev\\003_study\\docxsample\\fop.xconf";

        
        
		/**
		1.split docx
		2.docx to pdf
		3.pdf merge
		*/
        
        long startTime = System.currentTimeMillis(); // 코드 시작 시간
        
        try {
        	
            /**********************************************************************************
             * 1. docx 분할작업.
             **********************************************************************************/
        	
            // 사용 예제: 파일당 1000개 단락씩 자동 분할
            List<String> files = DocxSplitter.splitDocx(inputfilepath, mainfilepath, 100);
            
            System.out.println("\n=== 생성된 파일 목록 ===");
            
            /**********************************************************************************
             * 2. docx to pdf
             **********************************************************************************/
            List<File> fileList = new ArrayList<>(); // List<File> 생성
            
            docx2pdf con = new docx2pdf();
            
            for (int i = 0; i < files.size(); i++) {
                System.out.println((i + 1) + ". " + files.get(i));
                String nmpath = files.get(i).substring(0, files.get(i).length()-4 );
            	File newFile=con.toPDF(files.get(i), nmpath+"pdf", fopConfigPath);
            	
            	fileList.add(newFile);
            }

            
            /**********************************************************************************
             * 3. pdf files to merge pdf
             **********************************************************************************/
            String outputfilepath = mainfilepath + "output_merge_"+"모의침투결과"+".pdf";
            File outputFile = new File(outputfilepath);
            mergePdf(fileList,outputFile);
            

            
            
            /**********************************************************************************
             * 4. 초기화. temp file 제거.
             **********************************************************************************/
            deleteFiles(files);
            deleteFileList(fileList);
            
            
            
        	long endTime = System.currentTimeMillis(); // 코드 끝난 시간
        	long durationTimeSec = endTime - startTime;
            
        	System.out.println(durationTimeSec + "m/s"); // 밀리세컨드 출력
        	System.out.println((durationTimeSec / 1000) + "sec"); // 초 단위 변환 출력
            
        } catch (Exception e) {
            System.err.println("오류 발생:");
            e.printStackTrace();
        }
	}
	
	
	public static void mergePdf(List<File> pdfs, File output) throws Exception {
	    PDFMergerUtility merger = new PDFMergerUtility();
	    merger.setDestinationFileName(output.getAbsolutePath());

	    for (File pdf : pdfs) {
	        merger.addSource(pdf);
	    }

	    merger.mergeDocuments(null);
	}
	
	
	public static void deleteFiles(List<String> files) {
        for (String filePath : files) {
            try {
                // 파일 삭제 (파일이 없으면 NoSuchFileException 발생)
                Files.delete(Paths.get(filePath));
                System.out.println("삭제 성공: " + filePath);
            } catch (IOException e) {
                System.err.println("삭제 실패 (" + filePath + "): " + e.getMessage());
            }
        }
    }
	
	public static void deleteFileList(List<File> fileList) {
	    for (File file : fileList) {
	        if (file.exists()) { // 파일이 존재하는지 확인
	            if (file.delete()) {
	                System.out.println("삭제 성공: " + file.getName());
	            } else {
	                System.out.println("삭제 실패: " + file.getName());
	            }
	        }
	    }
	}

}
