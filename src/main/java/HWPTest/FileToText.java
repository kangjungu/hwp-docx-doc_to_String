package HWPTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.StringWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

import com.argo.hwp.HwpTextExtractor;

import difflib.Delta;
import difflib.DiffUtils;
import difflib.Patch;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

/*
 * org.apache.poi 라이브러리
 * 아파치 소프트웨어 재단에서 만든 라이브러리로서 마이크로소프트 오피스 파일 포맷을 순수 자바 언어로서 읽고 쓰는 기능을 제공한다
 */

public class FileToText {

	private static String result;
	private static Hashtable<Integer, List<?>> hash;
	private static Hashtable<Integer, String> typeHash;

	/*
	 * 두개의 파일을 내용을 비교해주는 메소드
	 *  파라미터 1 : 이전 버전  파일 경로
	 *  파라미터 2 : 다음 버전 파일 경로 
	 */
	public static String compareText(String firstFile, String secondFile){
		String firstString = "";
		String secondString = "";
		if(firstFile.endsWith("hwp")){
			firstString = hwpFileContentParser(firstFile);
		}else if(firstFile.endsWith("docx")){
			firstString = docxFileContentParser(firstFile);
		}else if(firstFile.endsWith("doc")){
			firstString = docFileContentParser(firstFile);
		}else{
			return "error";
		}
		
		if(secondFile.endsWith("hwp")){
			secondString = hwpFileContentParser(secondFile);
		}else if(secondFile.endsWith("docx")){
			secondString = docxFileContentParser(secondFile);
		}else if(secondFile.endsWith("doc")){
			secondString = docFileContentParser(secondFile);
		}else{
			return "error";
		}
		
		String result = getDiff(firstString,secondString,"\n");
		
		return result;
	}
	
	/* 파일을 텍스트파일로 변환해서 저장해주는 메소드
	 * 파라미터 1 : 현재 파일 경로 및 이름 파라미터 2 : 저장할 경로 파라미터 3 : 저장할 파일 이름
	 */
	public static boolean fileTotxtFile(String currentPath, String savePath, String saveFilename) {
		if (currentPath.endsWith("hwp")) {
			result = hwpFileContentParser(currentPath);
			if (!errorTest(result)) {
				return false;
			}
			return stringToFile(savePath, saveFilename, result);
		} else if (currentPath.endsWith("docx")) {
			result = docxFileContentParser(currentPath);
			if (!errorTest(result)) {
				return false;
			}
			return stringToFile(savePath, saveFilename, result);
		} else if (currentPath.endsWith("doc")) {
			result = docFileContentParser(currentPath);
			if (!errorTest(result)) {
				return false;
			}
			return stringToFile(savePath, saveFilename, result);
		} else {
			return false;
		}
	}

	// hwp파일에서 내용을 가져온다.
	private static String hwpFileContentParser(String fileName) {
		File hwp = new File(fileName); // 텍스트를 추출할 HWP 파일
		Writer writer = new StringWriter(); // 추출된 텍스트를 출력할 버퍼
		try {
			// 파일로부터 텍스트 추출
			HwpTextExtractor.extract(hwp, writer);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return "-1";
		} catch (IOException e) {
			e.printStackTrace();
			return "-1";
		}
		String text = writer.toString(); // 추출된 텍스트

		return text;
	}

	// doc, xls, ppt에서 내용을 가져온다.
	private static String docFileContentParser(String fileName) {
		POIFSFileSystem fs = null;
		try {

			fs = new POIFSFileSystem(new FileInputStream(fileName));

			if (fileName.endsWith(".doc")) {
				HWPFDocument doc = new HWPFDocument(fs);
				WordExtractor we = new WordExtractor(doc);
				return we.getText();
			} else if (fileName.endsWith(".xls")) {
				ExcelExtractor ex = new ExcelExtractor(fs);
				ex.setFormulasNotResults(true);
				ex.setIncludeSheetNames(true);
				return ex.getText();
			} else if (fileName.endsWith(".ppt")) {
				PowerPointExtractor extractor = new PowerPointExtractor(fs);
				return extractor.getText();
			}

		} catch (Exception e) {
//			 System.out.println("document file cant be indexed");
			return "-1";
		}
		return "-1";
	}

	// docx, xlsx, pptx에서 내용을 가져온다.
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
			// System.out.println("# DocxFileParser Error :" + e.getMessage());
			return "-1";
		}
		return "-1";
	}

	// string객체를 파일로 만들어준다.
	private static boolean stringToFile(String path, String fileName, String content) {
		String url = path;
		if (!fileName.endsWith(".txt")) {
			fileName = fileName.concat(".txt");
		}
		if (path.endsWith("/")) {
			url = path + fileName;
		} else {
			url = path + "/" + fileName;
		}

		try {
			FileUtils.writeStringToFile(new File(url), content, "UTF-8");
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
	}

	//두개의 스트링을 비교해서 다른부분에 태그 붙여준다
	private static String getDiff(String firstFile, String secondFile, String splitValue) {
		
		List<String> original = new ArrayList<String>(Arrays.asList(firstFile.split(splitValue)));
		List<String> revised = new ArrayList<String>(Arrays.asList(secondFile.split(splitValue)));
		
		hash = new Hashtable<Integer, List<?>>();
		typeHash = new Hashtable<>();
		Patch patch = DiffUtils.diff(original, revised);
		// 각각의 hashTable에 타입과 리스트를 넣어준다.
		for (Delta delta : patch.getDeltas()) {
			hash.put(delta.getRevised().getPosition(), delta.getRevised().getLines());
			typeHash.put(delta.getRevised().getPosition(), typeChange(delta.getType().toString()));
		}

		try {
			List<String> result = (List<String>) patch.applyTo(original);

			if (!result.equals(revised)) {
				return null;
			}

			StringBuilder stringList = new StringBuilder();

			for (int i = 0; i < result.size(); i++) {
				String s = result.get(i);
				// 변경,삭제,삽입된 부분이 있는경우 앞부분을 보기 쉽게 바꿔준다.
				if (hash.containsKey(i)) {
					// 바뀐 문장의 길이만큼 태그를 붙여준다.
					List<String> re = (List<String>) hash.get(i);
					for (int j = 0; j < re.size(); j++) {
						String change = result.get(i + j);
						if (change.equals(re.get(j))) {
							// 타입을 붙여준다.
							result.set(i + j, typeHash.get(i) + change);
						}
					}
				}
				s = result.get(i);
				if (i != result.size() - 1) {
					stringList.append(s + splitValue);
				}else{
					stringList.append(s);
				}

			}

			String merge = String.valueOf(stringList);

			// 결과 보내줌
			return merge;
		} catch (Exception e) {
			e.printStackTrace();
		} 
		return null;
	}

	// 결과가 -1이면 에러
	private static boolean errorTest(String result) {
		if (result.equals("-1")) {
			return false;
		}
		return true;
	}

	private static String typeChange(String type) {
		if (type.equals("INSERT")) {
			return Type.ADD.getType() + " ";
		} else if (type.equals("DELETE")) {
			return Type.DELETE.getType() + " ";
		} else {
			return Type.CHAGNE.getType() + " ";
		}
	}
}

// 추가,삭제,변경 타입
enum Type {
	ADD("$add"), DELETE("$del"), CHAGNE("$cha");

	private String type;

	Type(String type) {
		this.type = type;
	}

	public String getType() {
		return this.type;
	}
}