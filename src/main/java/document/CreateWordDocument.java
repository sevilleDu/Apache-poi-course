package document;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class CreateWordDocument {
	
	public static void main(String[] args) throws IOException {
		
		
		XWPFDocument document = new XWPFDocument();
		String filePath = "C:\\Users\\Owner\\Documents\\My Word Documents - Apache POI\\This is my first Word Document.docx";
		FileOutputStream output = new FileOutputStream(filePath);
		document.write(output);
		output.close();
		
		System.out.println("The Word Document has been created successfully!");		
		
		
	}

}
