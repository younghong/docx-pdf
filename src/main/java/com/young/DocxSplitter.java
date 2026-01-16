package com.young;


import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class DocxSplitter {
    
    /**
     * DOCX 파일을 단락 개수 기준으로 여러 개로 자동 분할합니다.
     * @param inputFile 원본 DOCX 파일 경로
     * @param outputDir 출력 디렉토리 경로
     * @param paragraphsPerFile 파일당 단락 개수 (예: 1000)
     * @return 생성된 파일 목록
     */
    public static List<String> splitDocx(String inputFile, String outputDir, int paragraphsPerFile) 
            throws Exception {
        
        List<String> createdFiles = new ArrayList<>();
        
        // 원본 DOCX 파일 읽기
        XWPFDocument originDoc = new XWPFDocument(new FileInputStream(inputFile));
        List<XWPFParagraph> allParagraphs = new ArrayList<>(originDoc.getParagraphs());
        
        System.out.println("총 단락 수: " + allParagraphs.size());
        System.out.println("파일당 단락 개수: " + paragraphsPerFile);
        
        // 분할 범위 자동 계산
        List<int[]> ranges = new ArrayList<>();
        for (int i = 0; i < allParagraphs.size(); i += paragraphsPerFile) {
            int endIdx = Math.min(i + paragraphsPerFile, allParagraphs.size());
            ranges.add(new int[]{i, endIdx});
        }
        
        System.out.println("생성될 파일 수: " + ranges.size());
        
        // 각 범위별로 파일 생성
        for (int fileNum = 0; fileNum < ranges.size(); fileNum++) {
            int[] range = ranges.get(fileNum);
            int startIdx = range[0];
            int endIdx = range[1];
            
            XWPFDocument newDoc = new XWPFDocument();
            
            // 범위 내의 단락과 관련 테이블만 복사
            int paraCount = 0;
            for (Object elem : originDoc.getBodyElements()) {
                if (elem instanceof XWPFParagraph) {
                    XWPFParagraph para = (XWPFParagraph) elem;
                    
                    // 범위에 해당하는 단락만 복사
                    if (paraCount >= startIdx && paraCount < endIdx) {
                        XWPFParagraph newPara = newDoc.createParagraph();
                        copyParagraphWithFormatting(para, newPara);
                    }
                    paraCount++;
                    
                } else if (elem instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) elem;
                    
                    // 범위 내에 적어도 하나의 단락이 있으면 테이블 복사
                    if (startIdx < paraCount && paraCount < endIdx) {
                        XWPFTable newTable = newDoc.createTable();
                        copyTableWithFormatting(table, newTable);
                    }
                }
            }
            
            // 파일 저장
            String outputFile = String.format("%s/output_%d.docx", outputDir, fileNum + 1);
            newDoc.write(new FileOutputStream(outputFile));
            createdFiles.add(outputFile);
            
            System.out.println("파일 " + (fileNum + 1) + " 생성: " + outputFile 
                + " (단락: " + (endIdx - startIdx) + "개)");
            
            newDoc.close();
        }
        
        originDoc.close();
        
        System.out.println("\n=== 분할 완료! ===");
        System.out.println("총 " + createdFiles.size() + "개 파일 생성");
        
        return createdFiles;
    }
    
    /**
     * 단락을 서식과 함께 복사합니다.
     */
    private static void copyParagraphWithFormatting(XWPFParagraph source, XWPFParagraph target) {
        try {
            // CTP (Core Text Properties) 전체 복사
            CTP sourceCTP = source.getCTP();
            CTP targetCTP = target.getCTP();
            
            // 단락 속성 복사
            if (sourceCTP.getPPr() != null) {
                targetCTP.setPPr((CTPPr) sourceCTP.getPPr().copy());
            }
            
            // 기존 run 제거
            while (targetCTP.getRList().size() > 0) {
                targetCTP.removeR(0);
            }
            
            // Run 복사
            for (CTR sourceRun : sourceCTP.getRList()) {
                CTR newRun = targetCTP.addNewR();
                newRun.set(sourceRun.copy());
            }
            
        } catch (Exception e) {
            System.err.println("단락 복사 중 오류: " + e.getMessage());
        }
    }
    
    /**
     * 테이블을 서식과 함께 복사합니다.
     */
    private static void copyTableWithFormatting(XWPFTable source, XWPFTable target) {
        try {
            CTTbl sourceTbl = source.getCTTbl();
            CTTbl targetTbl = target.getCTTbl();
            
            // 테이블 속성 복사
            if (sourceTbl.getTblPr() != null) {
                targetTbl.setTblPr((CTTblPr) sourceTbl.getTblPr().copy());
            }
            
            // 행 제거 후 복사
            while (targetTbl.getTrList().size() > 0) {
                targetTbl.removeTr(0);
            }
            
            // 각 행 복사
            for (CTRow sourceRow : sourceTbl.getTrList()) {
                CTRow newRow = targetTbl.addNewTr();
                newRow.set(sourceRow.copy());
            }
            
        } catch (Exception e) {
            System.err.println("테이블 복사 중 오류: " + e.getMessage());
        }
    }
    
}