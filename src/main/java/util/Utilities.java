package util;

import java.math.BigInteger;
import java.util.List;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

public class Utilities {

	public XWPFRun addParagraph(XWPFDocument document, String text, int spaces) {

		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText(text);

		for (int i = 0; i < spaces; i++) {
			run.addBreak();
		}

		return run;
	}

	public XWPFRun addCustomParagraps(XWPFDocument document, String text, int spaces, boolean makeBold,
			String fontColor, int fontSize) {

		String fontFamily = "Calibri";

		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();

		run.setText(text);
		run.setBold(makeBold);
		run.setColor(fontColor);
		run.setFontSize(fontSize);
		run.setFontFamily(fontFamily);
		run.getParagraph().setSpacingAfter(0);

		for (int i = 0; i < spaces; i++) {
			run.addBreak();
		}

		return run;
	}

	public XWPFRun addCustomParagrapsAdvanced(XWPFDocument document, String text, int spaces, boolean makeBold,
			String fontColor, int fontSize, String alignment, boolean underline, boolean indentation) {

		String fontFamily = "Calibri";

		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();

		run.setText(text);
		run.setBold(makeBold);
		run.setColor(fontColor);
		run.setFontSize(fontSize);
		run.setFontFamily(fontFamily);
		run.getParagraph().setSpacingAfter(0);

		for (int i = 0; i < spaces; i++) {
			run.addBreak();
		}

		/** Set alignment **/

		if (alignment.equalsIgnoreCase("center")) {
			run.getParagraph().setAlignment(ParagraphAlignment.CENTER);
		} else if (alignment.equalsIgnoreCase("justified")) {
			run.getParagraph().setAlignment(ParagraphAlignment.BOTH);
		} else if (alignment.equalsIgnoreCase("left")) {
			run.getParagraph().setAlignment(ParagraphAlignment.LEFT);
		}

		/** Set underline **/
		if (underline) {
			run.setUnderline(UnderlinePatterns.SINGLE);
		}

		/** Set indentation **/
		if (indentation) {
			run.getParagraph().setIndentationLeft(200);
		}

		return run;
	}

	public BigInteger generateListNumbering(XWPFDocument document, BigInteger abstractNumberingId,
			String numberingType) {

		CTAbstractNum ctabtractNum = CTAbstractNum.Factory.newInstance();
		ctabtractNum.setAbstractNumId(abstractNumberingId);

		CTLvl ctlvl = ctabtractNum.addNewLvl();
		ctlvl.setIlvl(BigInteger.valueOf(0));

		if (numberingType.equalsIgnoreCase("decimal")) {
			ctlvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
		} else if (numberingType.equalsIgnoreCase("lowercase")) {
			ctlvl.addNewNumFmt().setVal(STNumberFormat.LOWER_LETTER);
		} else if (numberingType.equalsIgnoreCase("uppercase")) {
			ctlvl.addNewNumFmt().setVal(STNumberFormat.UPPER_LETTER);
		}

		ctlvl.addNewLvlText().setVal("%1.");
		ctlvl.addNewStart().setVal(BigInteger.valueOf(1));

		XWPFAbstractNum abstractNum = new XWPFAbstractNum(ctabtractNum);
		XWPFNumbering numbering = document.createNumbering();

		abstractNumberingId = numbering.addAbstractNum(abstractNum);
		BigInteger numberingId = numbering.addNum(abstractNumberingId);

		return numberingId;

	}
	
	
	public void addFooter(XWPFDocument document, String footerText) {
		
		XWPFParagraph paragraph;
		XWPFRun run;
		
        XWPFHeaderFooterPolicy headerFooterPolicy = document.createHeaderFooterPolicy();
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        
        paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        run = paragraph.createRun();
        run.setFontSize(10);
        run.setText(footerText);

	}
	
	
	public void addFooterAndPageNumber(XWPFDocument document, String footerText) {
		
		XWPFParagraph paragraphPageNum;
		XWPFParagraph paragraph;
		XWPFRun run;
		
        XWPFHeaderFooterPolicy headerFooterPolicy = document.createHeaderFooterPolicy();
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        
        /**Adding footer text/content **/
        paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingAfter(0);

        run = paragraph.createRun();
        run.setFontSize(10);
        run.setText(footerText);
        
        /** Adding page number**/
        paragraphPageNum = footer.createParagraph();
        paragraphPageNum.setAlignment(ParagraphAlignment.RIGHT);
        paragraphPageNum.setSpacingAfter(0);
        
        run = paragraphPageNum.createRun();
        run.setText("Page ");
        run.setFontSize(10);
        
       
        run = paragraphPageNum.createRun();
        run.getParagraph().getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
        run.setFontSize(10);
        
        run = paragraphPageNum.createRun();
        run.setText(" of ");
        paragraphPageNum.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");
        run.setFontSize(10);
        run.getParagraph().setAlignment(ParagraphAlignment.RIGHT);
        
       
	}
	

}
