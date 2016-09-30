package HWPTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.StringWriter;
import java.io.Writer;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

import com.argo.hwp.HwpTextExtractor;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
/*
 * org.apache.poi 라이브러리
 * 아파치 소프트웨어 재단에서 만든 라이브러리로서 마이크로소프트 오피스 파일 포맷을 순수 자바 언어로서 읽고 쓰는 기능을 제공한다
 * 
 */
public class FileToText {
	
	private static String result;
	
	/*
	 * 파라미터 1 : 현재 파일 경로 및 이름 
	 * 파라미터 2 : 저장할 경로
	 * 파라미터 3 : 저장할 파일 이름
	 */ 
	public static boolean fileTotxtFile(String currentPath,String savePath,String saveFilename){
		if(currentPath.endsWith("hwp")){
			result = hwpFileContentParser(currentPath);
			if(!errorTest(result)){
				return false;
			}
			return stringToFile(savePath,saveFilename,result);
		}else if(currentPath.endsWith("docx")){
			result = docxFileContentParser(currentPath);
			if(!errorTest(result)){
				return false;
			}
			return stringToFile(savePath,saveFilename,result); 
		}else if(currentPath.endsWith("doc")){
			result = docFileContentParser(currentPath);
			if(!errorTest(result)){
				return false;
			}
			return stringToFile(savePath,saveFilename,result);
		}else{
			return false;
		}
	}

	//hwp파일 파서
	private static String hwpFileContentParser(String fileName) {
		File hwp = new File(fileName); // 텍스트를 추출할 HWP 파일
		Writer writer = new StringWriter(); // 추출된 텍스트를 출력할 버퍼
		try {
			// 파일로부터 텍스트 추출
			HwpTextExtractor.extract(hwp, writer);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			
		} catch (IOException e) {
			e.printStackTrace();
			return "-1";
		} 
		String text = writer.toString(); // 추출된 텍스트

		return text;
	}

	// doc, xls, ppt 파서
	private static String docFileContentParser(String fileName) {
		POIFSFileSystem fs = null;
		try {

			fs = new POIFSFileSystem(new FileInputStream(fileName));

			if (fileName.endsWith(".doc")) {
				HWPFDocument doc = new HWPFDocument(fs);
				WordExtractor we = new WordExtractor(doc);
				return we.getText();
			} else if (fileName.endsWith(".xls")) {
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(fileName));
				ExcelExtractor ex = new ExcelExtractor(fs);
				ex.setFormulasNotResults(true);
				ex.setIncludeSheetNames(true);
				return ex.getText();
			} else if (fileName.endsWith(".ppt")) {
				PowerPointExtractor extractor = new PowerPointExtractor(fs);
				return extractor.getText();
			}

		} catch (Exception e) {
//			System.out.println("document file cant be indexed");
			return "-1";
		}
		return "-1";
	}

	// docx, xlsx, pptx 파서
	private static String docxFileContentParser(String fileName) {
		try {
			FileInputStream fs = new FileInputStream(new File(fileName));
			OPCPackage d = OPCPackage.open(fs);
			if (fileName.endsWith(".docx")) {
				XWPFWordExtractor xw = new XWPFWordExtractor(d);
				return xw.getText();
			} else if (fileName.endsWith(".pptx")) {
				XSLFPowerPointExtractor xp = new XSLFPowerPointExtractor(d);
				return xp.getText();
			} else if (fileName.endsWith(".xlsx")) {
				XSSFExcelExtractor xe = new XSSFExcelExtractor(d);
				xe.setFormulasNotResults(true);
				xe.setIncludeSheetNames(true);
				return xe.getText();
			}
		} catch (Exception e) {
//			System.out.println("# DocxFileParser Error :" + e.getMessage());
			return "-1";
		}
		return "-1";
	}
	
	private static boolean stringToFile(String path,String fileName,String content){
		String url = path;
		if(!fileName.endsWith(".txt")){
			fileName = fileName.concat(".txt");
		}
		if(path.endsWith("/")){
			url = path+fileName;
		}else{
			url = path+"/"+fileName;
		}
		
		try {
			FileUtils.writeStringToFile(new File(url), content, "UTF-8");
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
	}
	
	//결과가 -1이면 에러 
	private static boolean errorTest(String result){
		if(result.equals("-1")){
			return false;
		}
		return true;
	}
}

