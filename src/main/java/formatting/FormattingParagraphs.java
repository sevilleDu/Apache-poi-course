package formatting;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import util.Utilities;

public class FormattingParagraphs {

	public static void main(String[] args) throws IOException{
		
		
		Utilities util = new Utilities();
		
		XWPFDocument document = new XWPFDocument();
		String filePath = System.getenv("USERPROFILE")  + "\\Documents\\My Word Documents - Apache POI\\Formatting Paragraphs.docx";
		File myFile = new File(filePath);
		
		util.addCustomParagraps(document, "Apache POI - Component Overview", 2, true, "008080", 16);
		
		util.addCustomParagraps(document, "The Apache POI project is the master project for developing pure Java ports of file formats based on Microsoft's OLE 2"+
				" Compound Document Format. OLE 2 Compound Document Format is used by Microsoft Office Documents, as well as by programs using MFC"+
				" property sets to serialize their document objects.", 1, false, "5D6D7E", 12);
		
		util.addCustomParagraps(document, "Apache POI is also the master project for developing pure Java ports of file formats based on Office Open XML (ooxml)."+
				" OOXML is part of an ECMA / ISO standardisation effort. This documentation is quite large, but you can normally find the bit you need"+
				" without too much effort! ECMA-376 standard is here, and is also under the Microsoft OSP.", 1, false, "5D6D7E", 12);
		
		
		util.addCustomParagraps(document, "POIFS is the oldest and most stable part of POI. It is our port of the OLE 2 Compound Document Format to pure Java."+
				" It supports both read and write functionality. All of our components for the binary (non-XML) Microsoft Office formats ultimately"+
				" rely on it by definition. Please see the POIFS project page for more information.", 1, false, "5D6D7E", 12);
		
		
		
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
