package com.young;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.apache.pdfbox.multipdf.PDFMergerUtility;

public class convert2 {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		
		String mainfilepath = "D:\\dev\\003_study\\docxsample\\merge\\";
		
		String filename1="output201_모의침투결과"+".pdf";
		String filename2="output202_모의침투결과"+".pdf";
		
		String outputfilepath = mainfilepath + "output_merge_"+"모의침투결과"+".pdf";
		
		File file1 = new File(mainfilepath+filename1);
		File file2 = new File(mainfilepath+filename2);
		
		File outputFile = new File(outputfilepath);
		
		List<File> fileList = new ArrayList<>(); // List<File> 생성
		fileList.add(file1);
		fileList.add(file2);
		
		try {
			mergePdf(fileList,outputFile);
		} catch (Exception e) {
			// TODO Auto-generated catch block
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

	
}
