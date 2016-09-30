package HWPTest;

import java.io.File;
import java.io.IOException;
import java.util.Scanner;

public class textComapre {
	public static void main(String[] args) throws IOException {
		
		String currentPath = "/Users/kangjungu1/Downloads/02.docx";
		String savePath = "/Users/kangjungu1/Downloads";
		String saveFilename = "5.txt";
		
		FileToText.fileTotxtFile(currentPath, savePath, saveFilename);
	}

}