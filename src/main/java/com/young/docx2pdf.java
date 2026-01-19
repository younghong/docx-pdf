package com.young;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashSet;
import java.util.Set;

import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.wml.CTSettings;
import org.docx4j.wml.STTblLayoutType;
import org.docx4j.wml.TcPrInner.GridSpan;






import java.io.ByteArrayOutputStream;
import java.io.File;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.common.PDMetadata;

import org.apache.xmpbox.XMPMetadata;
import org.apache.xmpbox.schema.DublinCoreSchema;
import org.apache.xmpbox.schema.XMPBasicSchema;

import org.apache.xmpbox.xml.XmpSerializer;



public class docx2pdf {

	
	/**
	 * docx를 pdf로 변환하는 함수.
	 * @param inputPath 입력 파일
	 * @param outputPath 출력 파일
	 * @param xconfPath 한글 설정 파일
	 * @author 김영홍
	 */
	public File toPDF(String inputPath, String outputPath , String xconfPath)
	{
		System.out.println("자동 배포 TEST");
		
		File newFile = null;
		
		OutputStream os = null;
        try {
            // 1. FOP 설정 파일 경로 설정 (매우 중요!)
            if (new File(xconfPath).exists()) {
                System.setProperty("org.apache.fop.configuration", xconfPath);
                System.out.println("✓ FOP 설정 파일 적용");
            } else {
                System.out.println("⚠ FOP 설정 파일을 찾을 수 없습니다: " + xconfPath);
                System.out.println("  아래의 fop.xconf 파일을 생성해주세요.");
            }

            // 2. 시스템 폰트 자동 탐색
            System.out.println("📝 시스템 폰트 탐색 중...");
            PhysicalFonts.discoverPhysicalFonts();
            System.out.println("✓ 시스템 폰트 탐색 완료");

            // 3. WordprocessingMLPackage 로드
            System.out.println("📖 DOCX 파일 로드 중...");
            WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(inputPath));
            System.out.println("✓ DOCX 파일 로드 완료");

            
            removeAndFixDuplicateIds(wordMLPackage);
            
            
            
            // 4. fontTable.xml이 없을 경우 기본 폰트 설정
            // fontTable.xml에 폰트가 정의되지 않은 경우, 맑은 고딕을 기본 폰트로 사용
            addDefaultFontToDocx(wordMLPackage);
            System.out.println("✓ 기본 폰트(맑은 고딕) 설정 완료");
            
            
            // 4-1. 테이블 너비 자동 조정 (페이지 영역 초과 방지)
            adjustTableWidth(wordMLPackage);
            System.out.println("✓ 테이블 너비 조정 완료");

            // 5. 폰트 매퍼 설정
            Mapper fontMapper = new IdentityPlusMapper();
            wordMLPackage.setFontMapper(fontMapper);
            System.out.println("✓ 폰트 매퍼 설정 완료");
            System.out.println("✓ 폰트 매퍼 설정 완료");

            // 5. 출력 스트림 설정
            newFile=new File(outputPath);
            os = new FileOutputStream(newFile);

            // 6. DOCX를 PDF로 변환
            System.out.println("🔄 PDF로 변환 중...");
            Docx4J.toPDF(wordMLPackage, os);

            
            
            rewritePdfMetadata(newFile);
            
            
            System.out.println("\n✅ DOCX 파일이 PDF로 성공적으로 변환되었습니다.");
            System.out.println("📄 생성된 파일: " + outputPath);

        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("\n❌ 변환 중 오류 발생: " + e.getMessage());
            System.err.println("\n✓ 해결 방법:");
            System.err.println("  1. fop.xconf 파일을 설정했는지 확인");
            System.err.println("  2. 폰트 파일 경로가 올바른지 확인");
            System.err.println("  3. docx4j 버전을 최신으로 업데이트");
            System.err.println("  4. Maven 의존성 확인: docx4j-core, docx4j-export-fo");
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return newFile;
	}
	
	
	// 아래는 새로운 메서드 - 클래스에 추가
	private void fixAllBorderValues(WordprocessingMLPackage wordMLPackage) {
	    try {
	        org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
	        if (doc.getBody() != null) {
	            fixAllBorderValuesRecursive(doc.getBody());
	        }
	    } catch (Exception e) {
	        System.out.println("⚠ 테두리 값 수정 중 오류: " + e.getMessage());
	        e.printStackTrace();
	    }
	}

	private void fixAllBorderValuesRecursive(Object obj) {
	    if (obj == null) return;

	    // Body
	    if (obj instanceof org.docx4j.wml.Body) {
	        org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
	        for (Object child : body.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }

	    // JAXBElement 처리
	    if (obj instanceof javax.xml.bind.JAXBElement) {
	        javax.xml.bind.JAXBElement jaxbElement = (javax.xml.bind.JAXBElement) obj;
	        fixAllBorderValuesRecursive(jaxbElement.getValue());
	        return;
	    }

	    // Paragraph 처리
	    if (obj instanceof org.docx4j.wml.P) {
	        org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
	        org.docx4j.wml.PPr pPr = p.getPPr();
	        
	        if (pPr != null) {
	            // 문단 테두리 처리 - reflection 사용
	            fixParagraphBorderValues(pPr);
	        }
	        
	        for (Object child : p.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }

	    // Table 처리
	    if (obj instanceof org.docx4j.wml.Tbl) {
	        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
	        org.docx4j.wml.TblPr tblPr = tbl.getTblPr();
	        
	        if (tblPr != null) {
	            // 테이블 테두리 처리
	            org.docx4j.wml.TblBorders tblBorders = tblPr.getTblBorders();
	            if (tblBorders != null) {
	                fixBorderVal(tblBorders.getTop());
	                fixBorderVal(tblBorders.getLeft());
	                fixBorderVal(tblBorders.getBottom());
	                fixBorderVal(tblBorders.getRight());
	                fixBorderVal(tblBorders.getInsideH());
	                fixBorderVal(tblBorders.getInsideV());
	            }
	        }
	        
	        // 테이블 행과 셀 처리
	        for (Object child : tbl.getContent()) {
	            if (child instanceof org.docx4j.wml.Tr) {
	                org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) child;
	                for (Object trChild : tr.getContent()) {
	                    if (trChild instanceof javax.xml.bind.JAXBElement) {
	                        javax.xml.bind.JAXBElement jaxbEl = (javax.xml.bind.JAXBElement) trChild;
	                        Object tcObj = jaxbEl.getValue();
	                        
	                        if (tcObj instanceof org.docx4j.wml.Tc) {
	                            org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) tcObj;
	                            org.docx4j.wml.TcPr tcPr = tc.getTcPr();
	                            
	                            if (tcPr != null) {
	                                fixCellBorderValues(tcPr);
	                            }
	                            
	                            // 셀 내의 컨텐츠도 처리
	                            for (Object tcChild : tc.getContent()) {
	                                fixAllBorderValuesRecursive(tcChild);
	                            }
	                        }
	                    }
	                }
	            }
	        }
	        
	        // 재귀적으로 테이블 내용 처리
	        for (Object child : tbl.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }

	    // Run 처리
	    if (obj instanceof org.docx4j.wml.R) {
	        org.docx4j.wml.R r = (org.docx4j.wml.R) obj;
	        for (Object child : r.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }
	}

	// 테두리 val 속성 확인 및 설정
	private void fixBorderVal(org.docx4j.wml.CTBorder border) {
	    if (border != null) {
	        try {
	            if (border.getVal() == null) {
	                // val이 없으면 "single"로 기본값 설정
	                border.setVal(org.docx4j.wml.STBorder.SINGLE);
	            }
	        } catch (Exception e) {
	            // 예외 무시
	        }
	    }
	}

	// 문단 테두리 처리 - reflection 사용
	private void fixParagraphBorderValues(org.docx4j.wml.PPr pPr) {
	    try {
	        java.lang.reflect.Field[] fields = pPr.getClass().getDeclaredFields();
	        
	        for (java.lang.reflect.Field field : fields) {
	            field.setAccessible(true);
	            Object fieldValue = field.get(pPr);
	            
	            // CTBorder 타입 확인
	            if (fieldValue instanceof org.docx4j.wml.CTBorder) {
	                org.docx4j.wml.CTBorder border = (org.docx4j.wml.CTBorder) fieldValue;
	                if (border.getVal() == null) {
	                    border.setVal(org.docx4j.wml.STBorder.SINGLE);
	                }
	            }
	        }
	    } catch (Exception e) {
	        // 예외 무시
	    }
	}

	// 셀 테두리 처리
	private void fixCellBorderValues(org.docx4j.wml.TcPr tcPr) {
	    try {
	        // reflection을 사용하여 TcPr의 모든 필드 확인
	        java.lang.reflect.Field[] fields = tcPr.getClass().getDeclaredFields();
	        
	        for (java.lang.reflect.Field field : fields) {
	            field.setAccessible(true);
	            Object fieldValue = field.get(tcPr);
	            
	            // CTBorder 타입 확인
	            if (fieldValue instanceof org.docx4j.wml.CTBorder) {
	                org.docx4j.wml.CTBorder border = (org.docx4j.wml.CTBorder) fieldValue;
	                if (border.getVal() == null) {
	                    border.setVal(org.docx4j.wml.STBorder.SINGLE);
	                }
	            }
	        }
	    } catch (Exception e) {
	        // 예외 무시
	    }
	}
	
	
	private void rewritePdfMetadata(File pdfFile) throws Exception {

	    PDDocument doc = PDDocument.load(pdfFile);

	    // 1. Info Dictionary
	    PDDocumentInformation info = doc.getDocumentInformation();
	    info.setProducer("K PDF Engine");
	    info.setCreator("K DOCX Converter");
	    info.setTitle(pdfFile.getName());
	    info.setAuthor("MySystem");
	    doc.setDocumentInformation(info);

	    // 2. XMP
	    XMPMetadata xmp = XMPMetadata.createXMPMetadata();

	    XMPBasicSchema basic = xmp.createAndAddXMPBasicSchema();
	    basic.setCreatorTool("H PDF Engine");

	    DublinCoreSchema dc = xmp.createAndAddDublinCoreSchema();
	    dc.addCreator("My DOCX Converter");

	    PDMetadata metadata = new PDMetadata(doc);
	    ByteArrayOutputStream baos = new ByteArrayOutputStream();
	    new XmpSerializer().serialize(xmp, baos, true);
	    metadata.importXMPMetadata(baos.toByteArray());

	    doc.getDocumentCatalog().setMetadata(metadata);

	    doc.save(pdfFile);
	    doc.close();
	}
	
	
	
	
	
	
	
	
	
	
	// fontTable.xml이 없을 경우 기본 폰트를 설정하는 메서드
    private void addDefaultFontToDocx(WordprocessingMLPackage wordMLPackage) {
        try {
            org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
            DocumentSettingsPart settingsPart = wordMLPackage.getMainDocumentPart().getDocumentSettingsPart();
            
            if (settingsPart == null) {
                // DocumentSettingsPart가 없으면 생성
                settingsPart = new org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart();
                wordMLPackage.getMainDocumentPart().addTargetPart(settingsPart);
            }
            
            // 기본 폰트를 맑은 고딕으로 설정
            CTSettings settings = settingsPart.getContents();
            if (settings == null) {
                settings = new CTSettings();
                settingsPart.setContents(settings);
            }
            
            // ThemeFontScheme 설정 (기본 폰트 지정)
            org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
            org.docx4j.wml.RFonts rFonts = factory.createRFonts();
            rFonts.setAscii("맑은 고딕");
            rFonts.setHAnsi("맑은 고딕");
            rFonts.setCs("맑은 고딕");
            
            // 모든 문단과 텍스트에 기본 폰트 적용
            applyDefaultFontToAllElements(doc, rFonts);
            
        } catch (Exception e) {
            System.out.println("⚠ 기본 폰트 설정 중 오류: " + e.getMessage());
        }
    }

    private static final float A4_WIDTH_PX = 794f; // 96dpi 기준 A4 가로
    private static final float A4_PADDING_PX = 76f;
    
    
 // mm → px
    public static int mmToPx(double mm, double dpi) {
        return (int) Math.round(mm * dpi / 25.4);
    }
    
    
    // DXA → PX
    private static float dxaToPx(int dxa) {
        return dxa * 96f / 1440f;
    }
    
    public static int pxToDxa(int px) {
        return Math.round(px * 1440f / 96f);
    }

    // DXA 배열을 A4 가로폭에 맞게 px로 비율 축소
    public static int[] scaleToA4Px(int[] dxaArray) {
        float[] pxArray = new float[dxaArray.length];
        float totalPx = 0f;

        // 1. DXA → PX 변환
        for (int i = 0; i < dxaArray.length; i++) {
            pxArray[i] = dxaToPx(dxaArray[i]);
            totalPx += pxArray[i];
        }

        // 2. A4에 맞는 스케일 비율
        float scale = (A4_WIDTH_PX-(A4_PADDING_PX*2)) / totalPx;

        // 3. 비율 적용
        int[] result = new int[dxaArray.length];
        for (int i = 0; i < pxArray.length; i++) {
            result[i] = pxToDxa(Math.round(pxArray[i] * scale));
        }

        return result;
    }
    
    private void adjustTableWidth(WordprocessingMLPackage wordMLPackage) {
        try {
            org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
            org.docx4j.wml.Body body = doc.getBody();
            
            if (body != null) {
                for (Object bodyChild : body.getContent()) {
                    if (bodyChild instanceof javax.xml.bind.JAXBElement) {
                        javax.xml.bind.JAXBElement jaxbElement = (javax.xml.bind.JAXBElement)bodyChild;
                        Object tbltest = jaxbElement.getValue();
                        
                        if (tbltest instanceof org.docx4j.wml.Tbl) {
                            org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) tbltest;
                            
                            // 1. 각 행의 실제 열 개수 계산 (gridSpan 포함)
                            int maxColCount = 0;
                            for (Object tblChild : tbl.getContent()) {
                                if (tblChild instanceof org.docx4j.wml.Tr) {
                                    org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tblChild;
                                    int colCount = calculateActualColumnCount(tr);
                                    maxColCount = Math.max(maxColCount, colCount);
                                }
                            }
                            
                            // 2. TblGrid 수정 또는 생성
                            org.docx4j.wml.TblGrid tblGrid = tbl.getTblGrid();
                            if (tblGrid == null) {
                                tblGrid = new org.docx4j.wml.TblGrid();
                                tbl.setTblGrid(tblGrid);
                            }
                            
                            java.util.List<org.docx4j.wml.TblGridCol> gridCols = tblGrid.getGridCol();
                            
                            // gridCol 개수를 maxColCount에 맞추기
                            while (gridCols.size() < maxColCount) {
                                org.docx4j.wml.TblGridCol col = new org.docx4j.wml.TblGridCol();
                                col.setW(java.math.BigInteger.valueOf(1440));
                                gridCols.add(col);
                            }
                            
                            // 3. 각 행의 초과 셀 제거
                            for (Object tblChild : tbl.getContent()) {
                                if (tblChild instanceof org.docx4j.wml.Tr) {
                                    org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tblChild;
                                    removeExcessCells(tr, maxColCount);
                                }
                            }
                            
                            // 4. TblGrid 크기 조정
                            int[] dxaArray = new int[gridCols.size()];
                            for (int i = 0; i < gridCols.size(); i++) {
                                org.docx4j.wml.TblGridCol col = gridCols.get(i);
                                java.math.BigInteger w = col.getW();
                                dxaArray[i] = (w != null) ? w.intValue() : 1440;
                            }
                            
                            int[] dxaArrayResult = scaleToA4Px(dxaArray);
                            for (int i = 0; i < gridCols.size(); i++) {
                                gridCols.get(i).setW(java.math.BigInteger.valueOf(dxaArrayResult[i]));
                            }
                            
                            // 5. 테이블 속성 설정
                            org.docx4j.wml.TblPr tblPr = tbl.getTblPr();
                            if (tblPr == null) {
                                tblPr = new org.docx4j.wml.TblPr();
                                tbl.setTblPr(tblPr);
                            }
                            
                            org.docx4j.wml.TblWidth tblW = new org.docx4j.wml.TblWidth();
                            tblW.setW(java.math.BigInteger.valueOf(5000));
                            tblW.setType("pct");
                            tblPr.setTblW(tblW);
                            
                            org.docx4j.wml.CTTblLayoutType tblLayout = new org.docx4j.wml.CTTblLayoutType();
                            tblLayout.setType(org.docx4j.wml.STTblLayoutType.AUTOFIT);
                            tblPr.setTblLayout(tblLayout);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("⚠ 테이블 너비 조정 중 오류: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // gridSpan을 고려한 실제 열 개수 계산
    private int calculateActualColumnCount(org.docx4j.wml.Tr tr) {
        int colCount = 0;
        for (Object trChild : tr.getContent()) {
            if (trChild instanceof javax.xml.bind.JAXBElement) {
                javax.xml.bind.JAXBElement jaxbElementTc = (javax.xml.bind.JAXBElement)trChild;
                Object tCtest = jaxbElementTc.getValue();
                if (tCtest instanceof org.docx4j.wml.Tc) {
                    org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) tCtest;
                    org.docx4j.wml.TcPr tcPr = tc.getTcPr();
                    
                    int gridSpan = 1;
                    if (tcPr != null && tcPr.getGridSpan() != null) {
                        gridSpan = tcPr.getGridSpan().getVal().intValue();
                        
                    }
                    colCount += gridSpan;
                }
            }
        }
        return colCount;
    }

    // 초과 셀 제거 (gridSpan 고려)
    private void removeExcessCells(org.docx4j.wml.Tr tr, int maxColCount) {
        int currentColIndex = 0;
        java.util.List<Object> cellsToRemove = new java.util.ArrayList<>();
        
        for (Object trChild : tr.getContent()) {
            if (trChild instanceof javax.xml.bind.JAXBElement) {
                javax.xml.bind.JAXBElement jaxbElementTc = (javax.xml.bind.JAXBElement)trChild;
                Object tCtest = jaxbElementTc.getValue();
                if (tCtest instanceof org.docx4j.wml.Tc) {
                    org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) tCtest;
                    org.docx4j.wml.TcPr tcPr = tc.getTcPr();
                    
                    int gridSpan = 1;
                    if (tcPr != null && tcPr.getGridSpan() != null) {
                        gridSpan = tcPr.getGridSpan().getVal().intValue();
                    }
                    
                    // gridSpan을 maxColCount를 초과하지 않도록 조정
                    if (currentColIndex >= maxColCount) {
                        cellsToRemove.add(trChild);
                    } else if (currentColIndex + gridSpan > maxColCount) {
                        // gridSpan 줄이기
                        if (tcPr == null) {
                            tcPr = new org.docx4j.wml.TcPr();
                            tc.setTcPr(tcPr);
                        }
                        int newGridSpan = maxColCount - currentColIndex;
                        
                        GridSpan gs = new  GridSpan();
                        gs.setVal(java.math.BigInteger.valueOf(newGridSpan));
                        tcPr.setGridSpan(gs);
                        
                        
                        currentColIndex = maxColCount;
                    } else {
                        currentColIndex += gridSpan;
                    }
                }
            }
        }
        
        // 초과 셀 제거
        for (Object cellToRemove : cellsToRemove) {
            tr.getContent().remove(cellToRemove);
        }
    }
    private void adjustTableWidth2(WordprocessingMLPackage wordMLPackage) {
        try {
            org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
            org.docx4j.wml.Body body = doc.getBody();
            
            if (body != null) {
                for (Object bodyChild : body.getContent()) {
                	
                	
                    if (bodyChild instanceof javax.xml.bind.JAXBElement) {
                    	javax.xml.bind.JAXBElement jaxbElement = (javax.xml.bind.JAXBElement)bodyChild;
                    	Object tbltest=jaxbElement.getValue();
                    	
                    	 if (tbltest instanceof org.docx4j.wml.Tbl) {
                             org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) tbltest;
                             org.docx4j.wml.TblGrid tblGrid=tbl.getTblGrid();
                             
                             
                             if(tblGrid==null) {
                            	 
                            	// 각 행의 높이 자동 조정
                                 for (Object tblChild : tbl.getContent()) {
                                     if (tblChild instanceof org.docx4j.wml.Tr) {
                                         org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tblChild;
                                         
                                         
                                         
                                         
                                         org.docx4j.wml.TrPr trPr = tr.getTrPr();
                                         if (trPr == null) {
                                             trPr = new org.docx4j.wml.TrPr();
                                             tr.setTrPr(trPr);
                                         }
                                         
                                         // 현재 행의 셀 개수
                                         int cellCount = 0;
                                         java.util.List<Object> cellsToRemove = new java.util.ArrayList<>();

                                         
                                         // 각 셀의 너비 자동 조정
                                         for (Object trChild : tr.getContent()) {
                                        	 
                                        	 
                                        	 if(trChild instanceof javax.xml.bind.JAXBElement) {
                                            	 javax.xml.bind.JAXBElement jaxbElementTc = (javax.xml.bind.JAXBElement)trChild;
                                              	Object tCtest=jaxbElementTc.getValue();
                                              	
                                                if (tCtest instanceof org.docx4j.wml.Tc) {
                                                	
                                                    cellCount++;
                                                    // TblGrid 열 개수 초과 시 제거 대상 표시
                                                    //if (cellCount > tblGridCols.size()) {
                                                     //   cellsToRemove.add(trChild);
                                                    //}
                                                	
                                                    org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) tCtest;
                                                    org.docx4j.wml.TcPr tcPr = tc.getTcPr();
                                                    if (tcPr == null) {
                                                        tcPr = new org.docx4j.wml.TcPr();
                                                        tc.setTcPr(tcPr);
                                                    }
                                                    
                                                    // 셀 너비 제거 (테이블 자동 조정에 맡김)
                                                    tcPr.setTcW(null);
                                                }
                                        	 }
                                         }
                                     }
                                 }
                            	 
                            	 
                            	 
                            	 continue;
                             }
                             java.util.List<org.docx4j.wml.TblGridCol> tblGridCols = tblGrid.getGridCol();
                             
                             
                             int[] dxaArray = new int[tblGridCols.size()];

                             for (int i = 0; i < tblGridCols.size(); i++) {
                            	 org.docx4j.wml.TblGridCol col = tblGridCols.get(i);

                                 BigInteger w = col.getW();
                                 dxaArray[i] = (w != null) ? w.intValue() : 0;
                             }
                             int[] dxaArrayResult=scaleToA4Px(dxaArray);
                             
                             
                             for (int i = 0; i < tblGridCols.size(); i++) {
                            	 org.docx4j.wml.TblGridCol col = tblGridCols.get(i);
                            	 col.setW(BigInteger.valueOf(dxaArrayResult[i]));
                             }
                             
                             
                             // 테이블 속성 설정
                             org.docx4j.wml.TblPr tblPr = tbl.getTblPr();
                             if (tblPr == null) {
                                 tblPr = new org.docx4j.wml.TblPr();
                                 tbl.setTblPr(tblPr);
                             }
                             
                             // 테이블 너비를 100% (페이지 너비)로 설정
                             org.docx4j.wml.TblWidth tblW = new org.docx4j.wml.TblWidth();
                             tblW.setW(java.math.BigInteger.valueOf(5000)); // 페이지 너비의 약 100%
                             tblW.setType("pct");
                             tblPr.setTblW(tblW);
                             
                             // 테이블 레이아웃을 Auto로 설정 (셀 내용에 따라 자동 조정)
                             org.docx4j.wml.CTTblLayoutType tblLayout = new org.docx4j.wml.CTTblLayoutType();
                             
                             tblLayout.setType(STTblLayoutType.AUTOFIT);
//                             tblLayout.setType(STTblLayoutType.FIXED);
                             
                             //tblLayout.setType("auto");
                             tblPr.setTblLayout(tblLayout);
                             
                             // 각 행의 높이 자동 조정
                             for (Object tblChild : tbl.getContent()) {
                                 if (tblChild instanceof org.docx4j.wml.Tr) {
                                     org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tblChild;
                                     
                                     
                                     
                                     
                                     org.docx4j.wml.TrPr trPr = tr.getTrPr();
                                     if (trPr == null) {
                                         trPr = new org.docx4j.wml.TrPr();
                                         tr.setTrPr(trPr);
                                     }
                                     
                                     // 현재 행의 셀 개수
                                     int cellCount = 0;
                                     java.util.List<Object> cellsToRemove = new java.util.ArrayList<>();

                                     
                                     // 각 셀의 너비 자동 조정
                                     for (Object trChild : tr.getContent()) {
                                    	 
                                    	 
                                    	 if(trChild instanceof javax.xml.bind.JAXBElement) {
                                        	 javax.xml.bind.JAXBElement jaxbElementTc = (javax.xml.bind.JAXBElement)trChild;
                                          	Object tCtest=jaxbElementTc.getValue();
                                          	
                                            if (tCtest instanceof org.docx4j.wml.Tc) {
                                            	
                                                cellCount++;
                                                // TblGrid 열 개수 초과 시 제거 대상 표시
                                                if (cellCount > tblGridCols.size()) {
                                                    cellsToRemove.add(trChild);
                                                }
                                            	
                                                org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) tCtest;
                                                org.docx4j.wml.TcPr tcPr = tc.getTcPr();
                                                if (tcPr == null) {
                                                    tcPr = new org.docx4j.wml.TcPr();
                                                    tc.setTcPr(tcPr);
                                                }
                                                
                                                // 셀 너비 제거 (테이블 자동 조정에 맡김)
                                                tcPr.setTcW(null);
                                            }
                                    	 }
                                     }
                                     
                                     // 초과 셀 제거
                                     for (Object cellToRemove : cellsToRemove) {
                                         tr.getContent().remove(cellToRemove);
                                     }
                                 }
                             }
                             
                             //System.out.println("✓ 테이블 너비 조정됨");
                         }
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("⚠ 테이블 너비 조정 중 오류: " + e.getMessage());
        }
    }
    
    
    
    private void applyDefaultFontToAllElements(Object obj, org.docx4j.wml.RFonts defaultFont) {
        if (obj == null) return;
        
        // Document 처리 (최상위 객체)
        if (obj instanceof org.docx4j.wml.Document) {
            org.docx4j.wml.Document doc = (org.docx4j.wml.Document) obj;
            org.docx4j.wml.Body body = doc.getBody();
            if (body != null) {
                applyDefaultFontToAllElements(body, defaultFont);
            }
            return;
        }
        
        // Body 처리
        if (obj instanceof org.docx4j.wml.Body) {
            org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
            for (Object bodyChild : body.getContent()) {
                applyDefaultFontToAllElements(bodyChild, defaultFont);
            }
            return;
        }
        
        // 문단(P) 처리
        if (obj instanceof org.docx4j.wml.P) {
            org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
            java.util.List<Object> pContent = p.getContent();
            for (Object pChild : pContent) {
                applyDefaultFontToAllElements(pChild, defaultFont);
            }
            return;
        }
        
        // 테이블(Tbl) 처리
        if (obj instanceof org.docx4j.wml.Tbl) {
            org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
            for (Object tblChild : tbl.getContent()) {
                applyDefaultFontToAllElements(tblChild, defaultFont);
            }
            return;
        }
        
        // 테이블 행(Tr) 처리
        if (obj instanceof org.docx4j.wml.Tr) {
            org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) obj;
            for (Object trChild : tr.getContent()) {
                applyDefaultFontToAllElements(trChild, defaultFont);
            }
            return;
        }
        
        // 테이블 셀(Tc) 처리
        if (obj instanceof org.docx4j.wml.Tc) {
            org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) obj;
            for (Object tcChild : tc.getContent()) {
                applyDefaultFontToAllElements(tcChild, defaultFont);
            }
            return;
        }
        
        if (obj instanceof javax.xml.bind.JAXBElement) {
        	javax.xml.bind.JAXBElement jaxbElement = (javax.xml.bind.JAXBElement)obj;
        	Object tbltest=jaxbElement.getValue();
        	applyDefaultFontToAllElements(tbltest, defaultFont);
        	return;
        }
        
        
        // 텍스트 런(R) 처리 - 실제 폰트 적용
        if (obj instanceof org.docx4j.wml.R) {
            org.docx4j.wml.R r = (org.docx4j.wml.R) obj;
            org.docx4j.wml.RPr rPr = r.getRPr();
            if (rPr == null) {
                rPr = new org.docx4j.wml.RPr();
                r.setRPr(rPr);
            }
            org.docx4j.wml.RFonts rFonts = rPr.getRFonts();
            if (rFonts == null || (rFonts.getAscii() == null && rFonts.getHAnsi() == null)) {
                if (rFonts == null) {
                    rFonts = new org.docx4j.wml.RFonts();
                }
                rFonts.setAscii("맑은 고딕");
                rFonts.setHAnsi("맑은 고딕");
                rFonts.setCs("맑은 고딕");
                rPr.setRFonts(rFonts);
                //System.out.println("✓ 폰트 적용됨");
            }else if(rFonts.getHAnsi().equals("Times New Roman")) {
                rFonts.setAscii("맑은 고딕");
                rFonts.setHAnsi("맑은 고딕");
                rFonts.setCs("맑은 고딕");
                rPr.setRFonts(rFonts);
            }
            return;
        }
        
        //System.out.println("other class="+obj.getClass());
    }
    
    
    
    
    
    
    
    
    
	private void removeAndFixDuplicateIds(WordprocessingMLPackage wordMLPackage) {
		try {
			org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
			Set<Long> usedIds = new HashSet<>();
			Set<String> usedBookmarkNames = new HashSet<>();
			
			if (doc.getBody() != null) {
				removeAndFixDuplicateIdsRecursive(doc.getBody(), usedIds,usedBookmarkNames);
			}
			
		} catch (Exception e) {
			System.out.println("âš  ID ì¤'ë³µ ì œê±° ì¤' ì˜¤ë¥˜: " + e.getMessage());
		}
	}
    
    
    private void removeAndFixDuplicateIdsRecursive(Object obj, Set<Long> usedIds, Set<String> usedBookmarkNames) {
		if (obj == null) return;

		if (obj instanceof javax.xml.bind.JAXBElement) {
			javax.xml.bind.JAXBElement jaxbElement = (javax.xml.bind.JAXBElement) obj;
			removeAndFixDuplicateIdsRecursive(jaxbElement.getValue(), usedIds,usedBookmarkNames);
			return;
		}
		
		
		if (obj instanceof org.docx4j.wml.CTBookmark) {
			org.docx4j.wml.CTBookmark bookmarkStart = (org.docx4j.wml.CTBookmark) obj;
			try {
				BigInteger id = bookmarkStart.getId();
				
				String name = bookmarkStart.getName();
				
				if (name != null && !name.isEmpty()) {
					if (usedBookmarkNames.contains(name)) {
						bookmarkStart.setName(null);
					} else {
						usedBookmarkNames.add(name);
					}
				}
				
				if (id != null) {
					Long idValue = id.longValue();
					if (usedIds.contains(idValue)) {
						bookmarkStart.setId(null);
					} else {
						usedIds.add(idValue);
					}
				}
			} catch (Exception e) {
			}
			return;
		}
		

		// Body
		if (obj instanceof org.docx4j.wml.Body) {
			org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
			for (Object child : body.getContent()) {
				removeAndFixDuplicateIdsRecursive(child, usedIds,usedBookmarkNames);
			}
			return;
		}

		// Paragraph
		if (obj instanceof org.docx4j.wml.P) {
			org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
			for (Object child : p.getContent()) {
				removeAndFixDuplicateIdsRecursive(child, usedIds,usedBookmarkNames);
			}
			return;
		}

		// Run
		if (obj instanceof org.docx4j.wml.R) {
			org.docx4j.wml.R r = (org.docx4j.wml.R) obj;
			for (Object child : r.getContent()) {
				removeAndFixDuplicateIdsRecursive(child, usedIds,usedBookmarkNames);
			}
			return;
		}

		// Run Properties
		// RPr 처리 부분 - getElem() 대신 getContent() 사용
		// Run Properties
		if (obj instanceof org.docx4j.wml.RPr) {
		    org.docx4j.wml.RPr rPr = (org.docx4j.wml.RPr) obj;
		    try {
		        // PPr 객체의 모든 필드를 reflection으로 접근
		        java.lang.reflect.Field[] fields = rPr.getClass().getDeclaredFields();
		        for (java.lang.reflect.Field field : fields) {
		            field.setAccessible(true);
		            Object fieldValue = field.get(rPr);
		            if (fieldValue != null) {
		                removeAndFixDuplicateIdsRecursive(fieldValue, usedIds,usedBookmarkNames);
		            }
		        }
		    } catch (Exception e) {
		        // 필드 접근 실패 무시
		    }
		    return;
		}

		// Drawing
		if (obj instanceof org.docx4j.wml.Drawing) {
			org.docx4j.wml.Drawing drawing = (org.docx4j.wml.Drawing) obj;
			java.util.List<Object> drawingContent = drawing.getAnchorOrInline();
			if (drawingContent != null) {
				for (Object child : drawingContent) {
					removeAndFixDuplicateIdsRecursive(child, usedIds,usedBookmarkNames);
				}
			}
			return;
		}

		// Inline (인라인)
		if (obj instanceof org.docx4j.dml.wordprocessingDrawing.Inline) {
		    org.docx4j.dml.wordprocessingDrawing.Inline inline = 
		        (org.docx4j.dml.wordprocessingDrawing.Inline) obj;
		    try {
		        // docPr의 실제 타입을 확인하고 처리
		        Object docPr = inline.getDocPr();
		        if (docPr != null) {
		            // reflection을 사용해 안전하게 ID 접근
		            java.lang.reflect.Method getIdMethod = docPr.getClass().getMethod("getId");
		            Long id = (Long) getIdMethod.invoke(docPr);
		            
		            if (id != null && usedIds.contains(id)) {
		                java.lang.reflect.Method setIdMethod = docPr.getClass().getMethod("setId", Long.class);
		                setIdMethod.invoke(docPr, (Long) null);
		            } else if (id != null) {
		                usedIds.add(id);
		            }
		        }
		    } catch (Exception e) {
		        // 메서드 호출 실패 무시
		    }
		    return;
		}

		// Anchor (앵커)
		if (obj instanceof org.docx4j.dml.wordprocessingDrawing.Anchor) {
		    org.docx4j.dml.wordprocessingDrawing.Anchor anchor = 
		        (org.docx4j.dml.wordprocessingDrawing.Anchor) obj;
		    try {
		        Object docPr = anchor.getDocPr();
		        if (docPr != null) {
		            java.lang.reflect.Method getIdMethod = docPr.getClass().getMethod("getId");
		            Long id = (Long) getIdMethod.invoke(docPr);
		            
		            if (id != null && usedIds.contains(id)) {
		                java.lang.reflect.Method setIdMethod = docPr.getClass().getMethod("setId", Long.class);
		                setIdMethod.invoke(docPr, (Long) null);
		            } else if (id != null) {
		                usedIds.add(id);
		            }
		        }
		    } catch (Exception e) {
		        // 메서드 호출 실패 무시
		    }
		    return;
		}

		// Table
		if (obj instanceof org.docx4j.wml.Tbl) {
			org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
			for (Object child : tbl.getContent()) {
				removeAndFixDuplicateIdsRecursive(child, usedIds,usedBookmarkNames);
			}
			return;
		}

		// Table Row
		if (obj instanceof org.docx4j.wml.Tr) {
			org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) obj;
			for (Object child : tr.getContent()) {
				removeAndFixDuplicateIdsRecursive(child, usedIds,usedBookmarkNames);
			}
			return;
		}

		// Table Cell
		if (obj instanceof org.docx4j.wml.Tc) {
			org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) obj;
			for (Object child : tc.getContent()) {
				removeAndFixDuplicateIdsRecursive(child, usedIds,usedBookmarkNames);
			}
			return;
		}
	}
    
    
    
    

}
