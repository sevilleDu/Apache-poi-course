package document;

import java.io.File;
import java.io.IOException;

public class DeleteWordDocument {

	public static void main(String[] args) throws IOException{
		
		String filePath = System.getenv("USERPROFILE")  + "\\Documents\\My Word Documents - Apache POI\\This is my first Word Document.docx";
		File myFile = new File(filePath);
		
		
		if(myFile.exists()) {
			
			System.out.println(myFile.getName() + " exist in my computer!");
			myFile.delete();
			System.out.println(myFile.getName() + " has been deleted!");
			
			
		}

	}

}
