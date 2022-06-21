package paragraphs;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class AddingMultipleRunsToAParagraph {

	public static void main(String[] args) throws IOException{
		
		XWPFDocument document = new XWPFDocument();
		String filePath = System.getenv("USERPROFILE")  + "\\Documents\\My Word Documents - Apache POI\\Paragrah Document Two.docx";
		File myFile = new File(filePath);
		
		XWPFParagraph paragraph = document.createParagraph();
		
		XWPFRun run1 = paragraph.createRun();
		run1.setText("This is the text associated to run1 of paragraph 1. ");
		run1.setText("This is the second text associated to paragraph 1, run 1.");
		run1.addBreak();
		
		XWPFRun run2 = paragraph.createRun();
		run2.setText("This text is the text we added in run2 which belongs to paragraph 1.");
		
		List<XWPFParagraph> paragList = document.getParagraphs();
		System.out.println("The number of paragraphs is: " + paragList.size());
		
		paragList.forEach(s-> {
			System.out.println(s.getRuns());
		});
		
		paragList.forEach(s-> {
			System.out.println("Number of runs in list of paragraphs is: " + s.getRuns().size());
		});
		
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
