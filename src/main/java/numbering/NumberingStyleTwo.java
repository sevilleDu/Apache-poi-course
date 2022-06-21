package numbering;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import paragraph_list.ParagraphContentList;
import util.Utilities;

public class NumberingStyleTwo {
	
	private static ParagraphContentList paragrapList = new ParagraphContentList();

	public static void main(String[] args) throws IOException{
		
		
		Utilities util = new Utilities();
		
		String footerText = "Microsoft Word Automation in Java | Apache POI API";

		XWPFDocument document = new XWPFDocument();
		String filePath = System.getenv("USERPROFILE")
				+ "\\Documents\\My Word Documents - Apache POI\\Add footer and page number to Word Document.docx";
		File myFile = new File(filePath);
		
		/**Paragraphs of List One **/

		util.addCustomParagrapsAdvanced(document, paragrapList.headerListOne, 2, true, "008080", 16, "center", true,
				false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragOneListOne, 1, false, "000000", 12, "left",
				false, false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragTwoListOne, 1, false, "000000", 12, "left",
				false, false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragTheeListOne, 1, false, "000000", 12, "left",
				false, false);
		
		
		/**Paragraphs of List Two **/
		
		util.addCustomParagrapsAdvanced(document, paragrapList.headerListTwo, 2, true, "008080", 16, "center", true,
				false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragOneListTwo, 1, false, "000000", 12, "left",
				false, false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragTwoListTwo, 1, false, "000000", 12, "left",
				false, false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragThreeListTwo, 1, false, "000000", 12, "left",
				false, false);

		
		/**Add List Numbering **/
		addNumbering(document, util.generateListNumbering(document, BigInteger.valueOf(1), "lowercase"), 1, paragrapList.paragraphsSectionOne());
		addNumbering(document, util.generateListNumbering(document, BigInteger.valueOf(2), "uppercase"), 2, paragrapList.paragraphsSectionTwo());
		
		/**Add footer and page number **/
		util.addFooterAndPageNumber(document, footerText);

		FileOutputStream output = new FileOutputStream(filePath);
		document.write(output);
		output.close();

		if (myFile.exists()) {

			if (Desktop.isDesktopSupported()) {
				Desktop.getDesktop().open(myFile);
			}

		}


	}
	
	
	
	public static void addNumbering(XWPFDocument document, BigInteger numId, int listNumber, List<String> parags) {
		
		List<XWPFParagraph> list = document.getParagraphs();
		
		for (XWPFParagraph paragraph : list) {
			
			
			for (int i = 1; i < parags.size(); i++) {
				
				if(paragraph.getText().trim().contains(parags.get(i))) {
					
					paragraph.setNumID(BigInteger.valueOf(listNumber));
					paragraph.getCTP().getPPr().addNewRPr().addNewSz().setVal(BigInteger.valueOf(22));
					paragraph.setIndentFromLeft(400);
					paragraph.setIndentationHanging(400);
					
				}
				
			}
			
			
		}
		
	}

}
