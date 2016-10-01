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

//할일 del 부분 넣기, 이전꺼 text도 보내기 
public class FileToText {

	private static String result;
	private static Hashtable<Integer, List<?>> changeAdd;
	private static List<Integer> delete;
	private static Hashtable<Integer, String> typeHash;

	/*
	 * 두개의 파일을 내용을 비교해주고 byte[]로 보내주는 메소드 파라미터 1 : 이전 버전 파일 경로 파라미터 2 : 다음 버전 파일
	 * 경로
	 */
	public static byte[] compareTextToByte(String firstFile, String secondFile) {
		String firstString = "";
		String secondString = "";
		if (firstFile.endsWith("hwp")) {
			firstString = hwpFileContentParser(firstFile);
		} else if (firstFile.endsWith("docx")) {
			firstString = docxFileContentParser(firstFile);
		} else if (firstFile.endsWith("doc")) {
			firstString = docFileContentParser(firstFile);
		} else {
			return "error".getBytes();
		}

		if (secondFile.endsWith("hwp")) {
			secondString = hwpFileContentParser(secondFile);
		} else if (secondFile.endsWith("docx")) {
			secondString = docxFileContentParser(secondFile);
		} else if (secondFile.endsWith("doc")) {
			secondString = docFileContentParser(secondFile);
		} else {
			return "error".getBytes();
		}

		String result = getDiff(firstString, secondString, "\n");

		return result.getBytes();
	}

	/*
	 * 두개의 파일을 내용을 비교해주는 메소드 파라미터 1 : 이전 버전 파일 경로 파라미터 2 : 다음 버전 파일 경로
	 */
	public static String[] compareText(String firstFile, String secondFile) {
		String firstString = "";
		String secondString = "";
		String returnResult[] = new String[] { "error", "error" };
		if (firstFile.endsWith("hwp")) {
			firstString = hwpFileContentParser(firstFile);
		} else if (firstFile.endsWith("docx")) {
			firstString = docxFileContentParser(firstFile);
		} else if (firstFile.endsWith("doc")) {
			firstString = docFileContentParser(firstFile);
		} else {
			return returnResult;
		}

		if (secondFile.endsWith("hwp")) {
			secondString = hwpFileContentParser(secondFile);
		} else if (secondFile.endsWith("docx")) {
			secondString = docxFileContentParser(secondFile);
		} else if (secondFile.endsWith("doc")) {
			secondString = docFileContentParser(secondFile);
		} else {
			return returnResult;
		}

		String result = getDiff(firstString, secondString, "\n");
		returnResult[0] = firstString;
		returnResult[1] = result;
		return returnResult;
	}

	/*
	 * 파일을 텍스트파일로 변환해서 저장해주는 메소드 파라미터 1 : 현재 파일 경로 및 이름 파라미터 2 : 저장할 경로 파라미터 3 :
	 * 저장할 파일 이름
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
			// System.out.println("document file cant be indexed");
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

	// 두개의 스트링을 비교해서 다른부분에 태그 붙여준다
	private static String getDiff(String firstFile, String secondFile, String splitValue) {
		
		//원본 내용
		List<String> original = new ArrayList<String>(Arrays.asList(firstFile.split(splitValue)));
		//두번째 파일 내용
		List<String> revised = new ArrayList<String>(Arrays.asList(secondFile.split(splitValue)));
		
		//변경사항 저장할 변수들 초기화 
		changeAdd = new Hashtable<Integer, List<?>>();
		delete = new ArrayList<Integer>();
		typeHash = new Hashtable<>();
		
		//변경된 부분을 가져온다.
		Patch patch = DiffUtils.diff(original, revised);
		
		//변경된 부분을 가져와서 객체에 넣어준다.
		for (Delta delta : patch.getDeltas()) {
			//delete 타입인경우 delete 리스트에 넣어준다.
			if(delta.getType().toString().contains("DELETE")){
				delete.add(delta.getRevised().getPosition());
			}else{
				//나머지경우 chageAdd hash에 넣어준다.
				changeAdd.put(delta.getRevised().getPosition(), delta.getRevised().getLines());
			}
			//타입을 넣어준다.
			typeHash.put(delta.getRevised().getPosition(), typeChange(delta.getType().toString()));
		}

		try {
			//변경된 부분 리스트를 가져온다.
			List<String> result = (List<String>) patch.applyTo(original);

			if (!result.equals(revised)) {
				return null;
			}

			//결과를 저장할 리스트 선언 
			List<String> stringList = new ArrayList<>();

			//변경된 부분들에 태크를 붙여 저장한다.
			for (int i = 0; i < result.size(); i++) {
				String s = result.get(i);
				// 변경,삽입된 부분이 있는경우 태그를 붙여준다.
				if (changeAdd.containsKey(i)) {
					//변경된 부분을 가져온다.
					List<String> re = (List<String>) changeAdd.get(i);
					// 바뀐 부분이 붙어있으면 결과가 한줄로 나오기때문에 반복문을 돌면서 태그를 다 붙여준다.
					for (int j = 0; j < re.size(); j++) {
						String change = result.get(i + j);
						if (change.equals(re.get(j))) {
							result.set(i + j, typeHash.get(i) + change);
						}
					}
				}
				s = result.get(i);
				if (i != result.size() - 1) {
					stringList.add(s + splitValue);
				}else{
					stringList.add(s);
				}

			}

			//삭제된 부분의 타입을 list에 넣어준다.
			for(int i=0;i<delete.size();i++){
				//delete이면 (자기의 인덱스 -1 + position)의 자리에 $del 넣기
				int position = delete.get(i)+i+1;
				stringList.add(position-1, Type.DELETE.getType()+"\n");
			}
			
			StringBuilder merge = new StringBuilder();
			for(int i=0;i<stringList.size();i++){
				merge.append(stringList.get(i));
			}
			String returnResult = String.valueOf(merge);
			// 결과 보내줌
			return returnResult;
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
