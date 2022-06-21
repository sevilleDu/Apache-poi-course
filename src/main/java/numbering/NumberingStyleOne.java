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

public class NumberingStyleOne {

	private static ParagraphContentList paragrapList = new ParagraphContentList();

	public static void main(String[] args) throws IOException {

		Utilities util = new Utilities();

		XWPFDocument document = new XWPFDocument();
		String filePath = System.getenv("USERPROFILE")
				+ "\\Documents\\My Word Documents - Apache POI\\Numbering Style One.docx";
		File myFile = new File(filePath);

		util.addCustomParagrapsAdvanced(document, paragrapList.headerListOne, 2, true, "008080", 16, "center", true,
				false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragOneListOne, 1, false, "5D6D7E", 12, "justified",
				false, false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragTwoListOne, 1, false, "5D6D7E", 12, "justified",
				false, false);
		util.addCustomParagrapsAdvanced(document, paragrapList.paragTheeListOne, 1, false, "5D6D7E", 12, "justified",
				false, false);

		// Add numnbering style
		List<XWPFParagraph> list = document.getParagraphs();

		for (XWPFParagraph parag : list) {

			if (!parag.getText().trim().contains(paragrapList.headerListOne)) {

				parag.setNumID(BigInteger.valueOf(1));
				parag.setIndentFromLeft(400);
				parag.setIndentationHanging(400);
			}

		}

		FileOutputStream output = new FileOutputStream(filePath);
		document.write(output);
		output.close();

		if (myFile.exists()) {

			if (Desktop.isDesktopSupported()) {
				Desktop.getDesktop().open(myFile);
			}

		}

	}

}
