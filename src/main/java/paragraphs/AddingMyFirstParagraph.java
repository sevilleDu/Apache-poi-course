package paragraphs;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class AddingMyFirstParagraph {

	public static void main(String[] args) throws IOException {
		
		XWPFDocument document = new XWPFDocument();
		String filePath = System.getenv("USERPROFILE")  + "\\Documents\\My Word Documents - Apache POI\\Paragrah Document One.docx";
		File myFile = new File(filePath);
		
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		
		run.setText("This is my first text, which I am adding to this word document");
		
		FileOutputStream output = new FileOutputStream(filePath);
		document.write(output);
		output.close();
		
		if(myFile.exists()) {
			
			if(Desktop.isDesktopSupported()) {
				Desktop.getDesktop().open(myFile);
			}
			
		}
		

	}

}
