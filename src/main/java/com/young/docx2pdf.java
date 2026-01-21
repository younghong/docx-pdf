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
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STTblLayoutType;
import org.docx4j.wml.TcPrInner.GridSpan;

import jakarta.xml.bind.JAXBElement;

import java.io.ByteArrayOutputStream;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.common.PDMetadata;
import org.apache.xmpbox.XMPMetadata;
import org.apache.xmpbox.schema.DublinCoreSchema;
import org.apache.xmpbox.schema.XMPBasicSchema;
import org.apache.xmpbox.xml.XmpSerializer;

public class docx2pdf {

	/**
	 * docxを pdfに変換する関数.
	 * @param inputPath 入力ファイル
	 * @param outputPath 出力ファイル
	 * @param xconfPath ハングル設定ファイル
	 * @author 김영화
	 */
	public File toPDF(String inputPath, String outputPath , String xconfPath)
	{
		File newFile = null;
		
		OutputStream os = null;
        try {
            // 1. FOP設定ファイルパスを設定 (非常に重要!)
            if (new File(xconfPath).exists()) {
                System.setProperty("org.apache.fop.configuration", xconfPath);
                System.out.println("✓ FOP設定ファイルが適用されました");
            } else {
                System.out.println("⚠  FOP設定ファイルが見つかりません: " + xconfPath);
                System.out.println("  以下のfop.xconfファイルを作成してください.");
            }

         // [추가: FOP 설정 초기화]
            initializeFopConfiguration();
            
            
            // 2. システムフォント自動検索
            System.out.println("🔍 システムフォント検索中...");
            PhysicalFonts.discoverPhysicalFonts();
            System.out.println("✓ システムフォント検索完了");

            // 3. WordprocessingMLPackageをロード
            System.out.println("📄 DOCXファイルロード中...");
            WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(inputPath));
            System.out.println("✓ DOCXファイルロード完了");

            System.setProperty("docx4j.convert.out.pdf.viaXSLFO.lineHeightFix", "true");

            removeAndFixDuplicateIds(wordMLPackage);
            
            // 変換前にすべての段落のline spacingをexactly pt値に変える
//            for (Object o : wordMLPackage.getMainDocumentPart().getJAXBNodesViaXPath("//w:p", true)) {
//                P p = (P)o;
//                PPr pPr = p.getPPr();
//                if (pPr == null) {
//                    pPr = new PPr();
//                    p.setPPr(pPr);
//                }
//                Spacing spacing = pPr.getSpacing();
//                if (spacing == null) {
//                    spacing = new Spacing();
//                    pPr.setSpacing(spacing);
//                }
//                // 例: 240 = 12pt exactly
//                spacing.setLineRule(STLineSpacingRule.EXACT);
//                spacing.setLine(BigInteger.valueOf(480));  // 希望するpt × 20
//            }
            
            preserveLineSpacingAndEmptyParagraphs(wordMLPackage);
            
            
            
         // 위치: Docx4J.toPDF(wordMLPackage, os); 호출 직전

         // [추가 코드 시작]
         applyLineSpacingToAllParagraphs(wordMLPackage);
         // [추가 코드 끝]
            

            // 5. フォントマッパー設定
            System.out.println("✓ フォントマッパー設定完了");
            System.out.println("✓ フォントマッパー設定完了");

            // 5. 出力ストリーム設定
            newFile=new File(outputPath);
            os = new FileOutputStream(newFile);

            // 6. DOCXをPDFに変換
            System.out.println("📄 PDFに変換中...");
            Docx4J.toPDF(wordMLPackage, os);

            rewritePdfMetadata(newFile);
            
            System.out.println("\n✅ DOCXファイルがPDFに正常に変換されました.");
            System.out.println("📄 生成されたファイル: " + outputPath);

        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("\n❌ 変換中にエラーが発生しました: " + e.getMessage());
            System.err.println("\n✓ 解決方法:");
            System.err.println("  1. fop.xconfファイルを設定したか確認");
            System.err.println("  2. フォントファイルパスが正しいか確認");
            System.err.println("  3. docx4jバージョンを最新にアップデート");
            System.err.println("  4. Maven依存性確認: docx4j-core, docx4j-export-fo");
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
	
	private void fixAllBorderValues(WordprocessingMLPackage wordMLPackage) {
	    try {
	        org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
	        if (doc.getBody() != null) {
	            fixAllBorderValuesRecursive(doc.getBody());
	        }
	    } catch (Exception e) {
	        System.out.println("⚠  テーダリー値更新中のエラー: " + e.getMessage());
	        e.printStackTrace();
	    }
	}

	private void fixAllBorderValuesRecursive(Object obj) {
	    if (obj == null) return;

	    if (obj instanceof org.docx4j.wml.Body) {
	        org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
	        for (Object child : body.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof JAXBElement) {
	        JAXBElement jaxbElement = (JAXBElement) obj;
	        fixAllBorderValuesRecursive(jaxbElement.getValue());
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.P) {
	        org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
	        org.docx4j.wml.PPr pPr = p.getPPr();
	        
	        if (pPr != null) {
	            fixParagraphBorderValues(pPr);
	        }
	        
	        for (Object child : p.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tbl) {
	        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
	        org.docx4j.wml.TblPr tblPr = tbl.getTblPr();
	        
	        if (tblPr != null) {
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
	        
	        for (Object child : tbl.getContent()) {
	            if (child instanceof org.docx4j.wml.Tr) {
	                org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) child;
	                for (Object trChild : tr.getContent()) {
	                    if (trChild instanceof JAXBElement) {
	                        JAXBElement jaxbEl = (JAXBElement) trChild;
	                        Object tcObj = jaxbEl.getValue();
	                        
	                        if (tcObj instanceof org.docx4j.wml.Tc) {
	                            org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) tcObj;
	                            org.docx4j.wml.TcPr tcPr = tc.getTcPr();
	                            
	                            if (tcPr != null) {
	                                fixCellBorderValues(tcPr);
	                            }
	                            
	                            for (Object tcChild : tc.getContent()) {
	                                fixAllBorderValuesRecursive(tcChild);
	                            }
	                        }
	                    }
	                }
	            }
	        }
	        
	        for (Object child : tbl.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.R) {
	        org.docx4j.wml.R r = (org.docx4j.wml.R) obj;
	        for (Object child : r.getContent()) {
	            fixAllBorderValuesRecursive(child);
	        }
	        return;
	    }
	}

	private void fixBorderVal(org.docx4j.wml.CTBorder border) {
	    if (border != null) {
	        try {
	            if (border.getVal() == null) {
	                border.setVal(org.docx4j.wml.STBorder.SINGLE);
	            }
	        } catch (Exception e) {
	            // 例外無視
	        }
	    }
	}

	private void fixParagraphBorderValues(org.docx4j.wml.PPr pPr) {
	    try {
	        java.lang.reflect.Field[] fields = pPr.getClass().getDeclaredFields();
	        
	        for (java.lang.reflect.Field field : fields) {
	            field.setAccessible(true);
	            Object fieldValue = field.get(pPr);
	            
	            if (fieldValue instanceof org.docx4j.wml.CTBorder) {
	                org.docx4j.wml.CTBorder border = (org.docx4j.wml.CTBorder) fieldValue;
	                if (border.getVal() == null) {
	                    border.setVal(org.docx4j.wml.STBorder.SINGLE);
	                }
	            }
	        }
	    } catch (Exception e) {
	        // 例外無視
	    }
	}

	private void fixCellBorderValues(org.docx4j.wml.TcPr tcPr) {
	    try {
	        java.lang.reflect.Field[] fields = tcPr.getClass().getDeclaredFields();
	        
	        for (java.lang.reflect.Field field : fields) {
	            field.setAccessible(true);
	            Object fieldValue = field.get(tcPr);
	            
	            if (fieldValue instanceof org.docx4j.wml.CTBorder) {
	                org.docx4j.wml.CTBorder border = (org.docx4j.wml.CTBorder) fieldValue;
	                if (border.getVal() == null) {
	                    border.setVal(org.docx4j.wml.STBorder.SINGLE);
	                }
	            }
	        }
	    } catch (Exception e) {
	        // 例外無視
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
	
	private void addDefaultFontToDocx(WordprocessingMLPackage wordMLPackage) {
	    try {
	        org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
	        DocumentSettingsPart settingsPart = wordMLPackage.getMainDocumentPart().getDocumentSettingsPart();
	        
	        if (settingsPart == null) {
	            settingsPart = new org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart();
	            wordMLPackage.getMainDocumentPart().addTargetPart(settingsPart);
	        }
	        
	        CTSettings settings = settingsPart.getContents();
	        if (settings == null) {
	            settings = new CTSettings();
	            settingsPart.setContents(settings);
	        }
	        
	        org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
	        org.docx4j.wml.RFonts rFonts = factory.createRFonts();
	        rFonts.setAscii("맑은 고딕");
	        rFonts.setHAnsi("맑은 고딕");
	        rFonts.setCs("맑은 고딕");
	        
	        applyDefaultFontToAllElements(doc, rFonts);
	        
	//        preserveLineSpacingAndEmptyParagraphs(wordMLPackage);
	        System.out.println("✓ 行間と空段落の保存完了");
	        
	    } catch (Exception e) {
	        System.out.println("⚠  デフォルトフォント設定中のエラー: " + e.getMessage());
	    }
	}

	private static final float A4_WIDTH_PX = 794f;
	private static final float A4_PADDING_PX = 76f;
	
	public static int mmToPx(double mm, double dpi) {
	    return (int) Math.round(mm * dpi / 25.4);
	}
	
	private static float dxaToPx(int dxa) {
	    return dxa * 96f / 1440f;
	}
	
	public static int pxToDxa(int px) {
	    return Math.round(px * 1440f / 96f);
	}

	public static int[] scaleToA4Px(int[] dxaArray) {
	    float[] pxArray = new float[dxaArray.length];
	    float totalPx = 0f;

	    for (int i = 0; i < dxaArray.length; i++) {
	        pxArray[i] = dxaToPx(dxaArray[i]);
	        totalPx += pxArray[i];
	    }

	    float scale = (A4_WIDTH_PX-(A4_PADDING_PX*2)) / totalPx;

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
	                if (bodyChild instanceof JAXBElement) {
	                    JAXBElement jaxbElement = (JAXBElement)bodyChild;
	                    Object tbltest = jaxbElement.getValue();
	                    
	                    if (tbltest instanceof org.docx4j.wml.Tbl) {
	                        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) tbltest;
	                        
	                        int maxColCount = 0;
	                        for (Object tblChild : tbl.getContent()) {
	                            if (tblChild instanceof org.docx4j.wml.Tr) {
	                                org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tblChild;
	                                int colCount = calculateActualColumnCount(tr);
	                                maxColCount = Math.max(maxColCount, colCount);
	                            }
	                        }
	                        
	                        org.docx4j.wml.TblGrid tblGrid = tbl.getTblGrid();
	                        if (tblGrid == null) {
	                            tblGrid = new org.docx4j.wml.TblGrid();
	                            tbl.setTblGrid(tblGrid);
	                        }
	                        
	                        java.util.List<org.docx4j.wml.TblGridCol> gridCols = tblGrid.getGridCol();
	                        
	                        while (gridCols.size() < maxColCount) {
	                            org.docx4j.wml.TblGridCol col = new org.docx4j.wml.TblGridCol();
	                            col.setW(java.math.BigInteger.valueOf(1440));
	                            gridCols.add(col);
	                        }
	                        
	                        for (Object tblChild : tbl.getContent()) {
	                            if (tblChild instanceof org.docx4j.wml.Tr) {
	                                org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tblChild;
	                                removeExcessCells(tr, maxColCount);
	                            }
	                        }
	                        
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
	        System.out.println("⚠  テーダル幅調整中のエラー: " + e.getMessage());
	        e.printStackTrace();
	    }
	}

	private int calculateActualColumnCount(org.docx4j.wml.Tr tr) {
	    int colCount = 0;
	    for (Object trChild : tr.getContent()) {
	        if (trChild instanceof JAXBElement) {
	            JAXBElement jaxbElementTc = (JAXBElement)trChild;
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

	private void removeExcessCells(org.docx4j.wml.Tr tr, int maxColCount) {
	    int currentColIndex = 0;
	    java.util.List<Object> cellsToRemove = new java.util.ArrayList<>();
	    
	    for (Object trChild : tr.getContent()) {
	        if (trChild instanceof JAXBElement) {
	            JAXBElement jaxbElementTc = (JAXBElement)trChild;
	            Object tCtest = jaxbElementTc.getValue();
	            if (tCtest instanceof org.docx4j.wml.Tc) {
	                org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) tCtest;
	                org.docx4j.wml.TcPr tcPr = tc.getTcPr();
	                
	                int gridSpan = 1;
	                if (tcPr != null && tcPr.getGridSpan() != null) {
	                    gridSpan = tcPr.getGridSpan().getVal().intValue();
	                }
	                
	                if (currentColIndex >= maxColCount) {
	                    cellsToRemove.add(trChild);
	                } else if (currentColIndex + gridSpan > maxColCount) {
	                    if (tcPr == null) {
	                        tcPr = new org.docx4j.wml.TcPr();
	                        tc.setTcPr(tcPr);
	                    }
	                    int newGridSpan = maxColCount - currentColIndex;
	                    
	                    GridSpan gs = new GridSpan();
	                    gs.setVal(java.math.BigInteger.valueOf(newGridSpan));
	                    tcPr.setGridSpan(gs);
	                    
	                    currentColIndex = maxColCount;
	                } else {
	                    currentColIndex += gridSpan;
	                }
	            }
	        }
	    }
	    
	    for (Object cellToRemove : cellsToRemove) {
	        tr.getContent().remove(cellToRemove);
	    }
	}

	private void applyDefaultFontToAllElements(Object obj, org.docx4j.wml.RFonts defaultFont) {
	    if (obj == null) return;
	    
	    if (obj instanceof org.docx4j.wml.Document) {
	        org.docx4j.wml.Document doc = (org.docx4j.wml.Document) obj;
	        org.docx4j.wml.Body body = doc.getBody();
	        if (body != null) {
	            applyDefaultFontToAllElements(body, defaultFont);
	        }
	        return;
	    }
	    
	    if (obj instanceof org.docx4j.wml.Body) {
	        org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
	        for (Object bodyChild : body.getContent()) {
	            applyDefaultFontToAllElements(bodyChild, defaultFont);
	        }
	        return;
	    }
	    
	    if (obj instanceof org.docx4j.wml.P) {
	        org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
	        java.util.List<Object> pContent = p.getContent();
	        for (Object pChild : pContent) {
	            applyDefaultFontToAllElements(pChild, defaultFont);
	        }
	        return;
	    }
	    
	    if (obj instanceof org.docx4j.wml.Tbl) {
	        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
	        for (Object tblChild : tbl.getContent()) {
	            applyDefaultFontToAllElements(tblChild, defaultFont);
	        }
	        return;
	    }
	    
	    if (obj instanceof org.docx4j.wml.Tr) {
	        org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) obj;
	        for (Object trChild : tr.getContent()) {
	            applyDefaultFontToAllElements(trChild, defaultFont);
	        }
	        return;
	    }
	    
	    if (obj instanceof org.docx4j.wml.Tc) {
	        org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) obj;
	        for (Object tcChild : tc.getContent()) {
	            applyDefaultFontToAllElements(tcChild, defaultFont);
	        }
	        return;
	    }
	    
	    if (obj instanceof JAXBElement) {
	        JAXBElement jaxbElement = (JAXBElement)obj;
	        Object tbltest=jaxbElement.getValue();
	        applyDefaultFontToAllElements(tbltest, defaultFont);
	        return;
	    }
	    
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
	        } else if(rFonts.getHAnsi() != null && rFonts.getHAnsi().equals("Times New Roman")) {
	            rFonts.setAscii("맑은 고딕");
	            rFonts.setHAnsi("맑은 고딕");
	            rFonts.setCs("맑은 고딕");
	            rPr.setRFonts(rFonts);
	        }
	        return;
	    }


	}
	
	
	
	
	private void removeAndFixDuplicateIds(WordprocessingMLPackage wordMLPackage) {
	    try {
	        org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
	        Set<Long> usedIds = new HashSet<>();
	        Set<String> usedBookmarkNames = new HashSet<>();
	        
	        if (doc.getBody() != null) {
	            removeAndFixDuplicateIdsRecursive(doc.getBody(), usedIds, usedBookmarkNames);
	        }
	        
	    } catch (Exception e) {
	        System.out.println("⚠  ID複製修正中のエラー: " + e.getMessage());
	    }
	}
	
	private void removeAndFixDuplicateIdsRecursive(Object obj, Set<Long> usedIds, Set<String> usedBookmarkNames) {
	    if (obj == null) return;

	    if (obj instanceof JAXBElement) {
	        JAXBElement jaxbElement = (JAXBElement) obj;
	        removeAndFixDuplicateIdsRecursive(jaxbElement.getValue(), usedIds, usedBookmarkNames);
	        return;
	    }
	    
	    if (obj instanceof org.docx4j.wml.CTBookmark) {
	        org.docx4j.wml.CTBookmark bookmarkStart = (org.docx4j.wml.CTBookmark) obj;
	        try {
	            BigInteger id = bookmarkStart.getId();
	            String name = bookmarkStart.getName();
	            
	            if (name != null && !name.isEmpty()) {
	                if (usedBookmarkNames.contains(name)) {
	                    bookmarkStart.setName("");
	                } else {
	                    usedBookmarkNames.add(name);
	                }
	            }
	            
	            if (id != null) {
	                Long idValue = id.longValue();
	                if (usedIds.contains(idValue)) {
	                    bookmarkStart.setId(new BigInteger(""));
	                } else {
	                    usedIds.add(idValue);
	                }
	            }
	        } catch (Exception e) {
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Body) {
	        org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
	        for (Object child : body.getContent()) {
	            removeAndFixDuplicateIdsRecursive(child, usedIds, usedBookmarkNames);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.P) {
	        org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
	        for (Object child : p.getContent()) {
	            removeAndFixDuplicateIdsRecursive(child, usedIds, usedBookmarkNames);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.R) {
	        org.docx4j.wml.R r = (org.docx4j.wml.R) obj;
	        for (Object child : r.getContent()) {
	            removeAndFixDuplicateIdsRecursive(child, usedIds, usedBookmarkNames);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.RPr) {
	        org.docx4j.wml.RPr rPr = (org.docx4j.wml.RPr) obj;
	        try {
	            java.lang.reflect.Field[] fields = rPr.getClass().getDeclaredFields();
	            for (java.lang.reflect.Field field : fields) {
	                field.setAccessible(true);
	                Object fieldValue = field.get(rPr);
	                if (fieldValue != null) {
	                    removeAndFixDuplicateIdsRecursive(fieldValue, usedIds, usedBookmarkNames);
	                }
	            }
	        } catch (Exception e) {
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Drawing) {
	        org.docx4j.wml.Drawing drawing = (org.docx4j.wml.Drawing) obj;
	        java.util.List<Object> drawingContent = drawing.getAnchorOrInline();
	        if (drawingContent != null) {
	            for (Object child : drawingContent) {
	                removeAndFixDuplicateIdsRecursive(child, usedIds, usedBookmarkNames);
	            }
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.dml.wordprocessingDrawing.Inline) {
	        org.docx4j.dml.wordprocessingDrawing.Inline inline = 
	            (org.docx4j.dml.wordprocessingDrawing.Inline) obj;
	        try {
	            Object docPr = inline.getDocPr();
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
	        }
	        return;
	    }

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
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tbl) {
	        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
	        for (Object child : tbl.getContent()) {
	            removeAndFixDuplicateIdsRecursive(child, usedIds, usedBookmarkNames);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tr) {
	        org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) obj;
	        for (Object child : tr.getContent()) {
	            removeAndFixDuplicateIdsRecursive(child, usedIds, usedBookmarkNames);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tc) {
	        org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) obj;
	        for (Object child : tc.getContent()) {
	            removeAndFixDuplicateIdsRecursive(child, usedIds, usedBookmarkNames);
	        }
	        return;
	    }
	}
	
	
	
	
	
	
	// ==========================================
	// 새로운 메서드 추가 (클래스 내부)
	// ==========================================

	private void applyLineSpacingToAllParagraphs(WordprocessingMLPackage wordMLPackage) {
	    try {
	        org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
	        if (doc.getBody() != null) {
	            applyLineSpacingRecursive(doc.getBody());
	        }
	    } catch (Exception e) {
	        System.out.println("⚠  줄간격 적용 중 오류: " + e.getMessage());
	        e.printStackTrace();
	    }
	}

	private void applyLineSpacingRecursive(Object obj) {
	    if (obj == null) return;

	    if (obj instanceof org.docx4j.wml.Body) {
	        org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
	        for (Object child : body.getContent()) {
	            applyLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof JAXBElement) {
	        JAXBElement jaxbElement = (JAXBElement) obj;
	        applyLineSpacingRecursive(jaxbElement.getValue());
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.P) {
	        org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
	        org.docx4j.wml.PPr pPr = p.getPPr();
	        
	        if (pPr != null) {
	            org.docx4j.wml.PPrBase.Spacing spacing = pPr.getSpacing();
	            
	            if (spacing != null && spacing.getLine() != null) {
	                BigInteger lineValue = spacing.getLine();
	                
	                // lineRule 설정
	                if (spacing.getLineRule() == null) {
	                    spacing.setLineRule(STLineSpacingRule.AUTO);
	                }
	                
	                // 줄간격 값 분석 및 로깅
	                double lineHeightPt = lineValue.doubleValue() / 20.0;
	                System.out.println("  - 적용된 줄간격: " + lineHeightPt + "pt (원본값: " + lineValue + ")");
	            }
	        }
	        
	        for (Object child : p.getContent()) {
	            applyLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tbl) {
	        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
	        for (Object child : tbl.getContent()) {
	            applyLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tr) {
	        org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) obj;
	        for (Object child : tr.getContent()) {
	            applyLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tc) {
	        org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) obj;
	        for (Object child : tc.getContent()) {
	            applyLineSpacingRecursive(child);
	        }
	        return;
	    }
	}
	
	
	// ==========================================
	// FOP 설정 초기화 메서드 (새로 추가)
	// ==========================================

	private void initializeFopConfiguration() {
	    try {
	        // FOP 라인 높이 수정 활성화
	        System.setProperty("docx4j.convert.out.pdf.viaXSLFO.lineHeightFix", "true");
	        
	        // FOP 설정 클래스 초기화
	        System.setProperty("org.apache.fop.dont-load-config-from-classpath", "true");
	        
	        // 글자 크기 기반 줄간격 계산 활성화
	        System.setProperty("docx4j.convert.out.pdf.viaXSLFO.lineHeightCorrection", "true");
	        
	        System.out.println("✓ FOP 설정 초기화 완료");
	        
	    } catch (Exception e) {
	        System.out.println("⚠  FOP 설정 초기화 중 오류: " + e.getMessage());
	    }
	}
	
	
	
	private void preserveLineSpacingAndEmptyParagraphs(WordprocessingMLPackage wordMLPackage) {
	    try {
	        org.docx4j.wml.Document doc = wordMLPackage.getMainDocumentPart().getContents();
	        if (doc.getBody() != null) {
	            preserveLineSpacingRecursive(doc.getBody());
	        }
	    } catch (Exception e) {
	        System.out.println("⚠  행간 보존 중 오류: " + e.getMessage());
	    }
	}

	private void preserveLineSpacingRecursive(Object obj) {
	    if (obj == null) return;

	    if (obj instanceof org.docx4j.wml.Body) {
	        org.docx4j.wml.Body body = (org.docx4j.wml.Body) obj;
	        for (Object child : body.getContent()) {
	            preserveLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof JAXBElement) {
	        JAXBElement jaxbElement = (JAXBElement) obj;
	        preserveLineSpacingRecursive(jaxbElement.getValue());
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.P) {
	        org.docx4j.wml.P p = (org.docx4j.wml.P) obj;
	        org.docx4j.wml.PPr pPr = p.getPPr();
	        
	        if (pPr == null) {
	            pPr = new org.docx4j.wml.PPr();
	            p.setPPr(pPr);
	        }
	        
	        // ============================================
	        // spacing 속성 확인 및 보정
	        // ============================================
	        org.docx4j.wml.PPrBase.Spacing spacing = pPr.getSpacing();
	        
	        if (spacing == null) {
	            spacing = new org.docx4j.wml.PPrBase.Spacing();
	            pPr.setSpacing(spacing);
	        }
	        
	        // w:before (단락 앞 공백)
	        if (spacing.getBefore() == null) {
	            spacing.setBefore(BigInteger.ZERO);
	        }
	        System.out.println("  - w:before: " + spacing.getBefore());
	        
	        // w:after (단락 뒤 공백) - 중요!
	        if (spacing.getAfter() == null) {
	            spacing.setAfter(BigInteger.ZERO);
	        }
	        System.out.println("  - w:after: " + spacing.getAfter());
	        
	        // w:line (줄간격)
	        if (spacing.getLine() != null) {
	            BigInteger lineValue = spacing.getLine();
	            
	            // lineRule 확인
	            if (spacing.getLineRule() == null) {
	                spacing.setLineRule(STLineSpacingRule.AUTO);
	            }
	            
	            System.out.println("  - w:line: " + lineValue + " (" + spacing.getLineRule() + ")");
	            
	            // 줄간격이 480 이상이면 AUTO 모드 유지
	            if (lineValue.compareTo(BigInteger.valueOf(480)) >= 0) {
	                //spacing.setLineRule(STLineSpacingRule.AUTO);
	                spacing.setLineRule(STLineSpacingRule.EXACT);
	                
	                
	                //spacing.setAfter(BigInteger.valueOf(480));
	                spacing.setAfter(spacing.getLine());
	                spacing.setBefore(spacing.getLine());
	                
	                System.out.println("  - 큰 줄간격 감지: AUTO 모드로 설정");
	            }
	        } else {
	            // w:line이 없으면 기본값 설정
	            spacing.setLine(BigInteger.valueOf(240));
	            spacing.setLineRule(STLineSpacingRule.AUTO);
	            System.out.println("  - 기본 줄간격 설정: 240 (AUTO)");
	        }
	        
	        // ============================================
	        // 단행 단락(빈 단락) 유지
	        // ============================================
	        if (p.getContent().isEmpty()) {
	            org.docx4j.wml.R r = new org.docx4j.wml.R();
	            p.getContent().add(r);
	        }
	        
	        for (Object child : p.getContent()) {
	            preserveLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tbl) {
	        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) obj;
	        for (Object child : tbl.getContent()) {
	            preserveLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tr) {
	        org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) obj;
	        for (Object child : tr.getContent()) {
	            preserveLineSpacingRecursive(child);
	        }
	        return;
	    }

	    if (obj instanceof org.docx4j.wml.Tc) {
	        org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) obj;
	        for (Object child : tc.getContent()) {
	            preserveLineSpacingRecursive(child);
	        }
	        return;
	    }
	}
}