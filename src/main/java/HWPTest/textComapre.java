package HWPTest;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.List;

import difflib.Delta;
import difflib.DiffUtils;
import difflib.Patch;
import difflib.PatchFailedException;

/*
 * 추가된경우 $add 
 * 삭제된경우 $del
 * 변경된경우 $cha 
 * 붙여서 한개로 만들어줌.
 * 
 */
public class textComapre {


	public static void main(String[] args) throws IOException {
		String result[] = FileToText.compareText("/Users/kangjungu1/Downloads/1.docx","/Users/kangjungu1/Downloads/2.docx");
		System.out.println(result[1]);
	}
	
}
