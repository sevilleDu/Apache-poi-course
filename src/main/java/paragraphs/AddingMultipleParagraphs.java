package paragraphs;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class AddingMultipleParagraphs {

	public static void main(String[] args) throws IOException {
		
		XWPFDocument document = new XWPFDocument();
		String filePath = System.getenv("USERPROFILE")  + "\\Documents\\My Word Documents - Apache POI\\Multiple paragraph Document.docx";
		File myFile = new File(filePath);
		
		XWPFParagraph paragraph1 = document.createParagraph();
		XWPFRun run1 = paragraph1.createRun();
		run1.setText("Apache POI - Component Overview");
		
		XWPFParagraph paragraph2 = document.createParagraph();
		XWPFRun run2 = paragraph2.createRun();
		run2.setText("The Apache POI project is the master project for developing pure Java ports of file formats based on Microsoft's OLE 2"+
		" Compound Document Format. OLE 2 Compound Document Format is used by Microsoft Office Documents, as well as by programs using MFC"+
		" property sets to serialize their document objects.");
		
		XWPFParagraph paragraph3 = document.createParagraph();
		XWPFRun run3 = paragraph3.createRun();
		run3.setText("Apache POI is also the master project for developing pure Java ports of file formats based on Office Open XML (ooxml)."+
		" OOXML is part of an ECMA / ISO standardisation effort. This documentation is quite large, but you can normally find the bit you need"+
		" without too much effort! ECMA-376 standard is here, and is also under the Microsoft OSP.");
		
		XWPFParagraph paragraph4 = document.createParagraph();
		XWPFRun run4 = paragraph4.createRun();
		run4.setText("POIFS is the oldest and most stable part of POI. It is our port of the OLE 2 Compound Document Format to pure Java."+
		" It supports both read and write functionality. All of our components for the binary (non-XML) Microsoft Office formats ultimately"+
		" rely on it by definition. Please see the POIFS project page for more information.");
		
		
		
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
