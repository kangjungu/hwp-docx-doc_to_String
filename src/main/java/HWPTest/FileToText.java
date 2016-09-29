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

import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

public class FileToText {
	
	public static String fileToText(String fileName){
		if(fileName.endsWith("hwp")){
			return hwpFileContentParser(fileName);
		}else if(fileName.endsWith("docx")){
			return docxFileContentParser(fileName);
		}else if(fileName.endsWith("doc")){
			return docFileContentParser(fileName);
		}else{
			return "";
		}
	}

	//hwp파일 파서
	private static String hwpFileContentParser(String fileName) {
		File hwp = new File(fileName); // 텍스트를 추출할 HWP 파일
		Writer writer = new StringWriter(); // 추출된 텍스트를 출력할 버퍼
		try {
			HwpTextExtractor.extract(hwp, writer);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} // 파일로부터 텍스트 추출
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
			System.out.println("document file cant be indexed");
		}
		return "";
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
			System.out.println("# DocxFileParser Error :" + e.getMessage());
		}
		return "";
	}
}
