package com.young;

import java.util.List;

public class App3 {

	public static void main(String[] args) {
		
		
		String mainfilepath = "D:\\dev\\003_study\\docxsample\\merge\\";
		String filename="한국가스공사";
        String inputfilepath = mainfilepath + filename+ ".docx";
        
        
        
        try {
            // 사용 예제: 파일당 1000개 단락씩 자동 분할
            List<String> files = DocxSplitter.splitDocx(inputfilepath, mainfilepath, 100);
            
            System.out.println("\n=== 생성된 파일 목록 ===");
            for (int i = 0; i < files.size(); i++) {
                System.out.println((i + 1) + ". " + files.get(i));
            }
            
        } catch (Exception e) {
            System.err.println("오류 발생:");
            e.printStackTrace();
        }
        
        
    }

}
