package document;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

public class OpenWordDocument {
	
	public static void main(String[] args) throws IOException {
		
		
		String filePath = System.getenv("USERPROFILE")  + "\\Documents\\My Word Documents - Apache POI\\This is my first Word Document.docx";
		File myFile = new File(filePath);
		
		if(myFile.exists()) {
			
			if(Desktop.isDesktopSupported()) {
				Desktop.getDesktop().open(myFile);
			}
		}
		
	}

}
