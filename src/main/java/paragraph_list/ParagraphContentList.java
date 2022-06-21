package paragraph_list;

import java.util.List;

public class ParagraphContentList {
	
	//List One
	
	public final String headerListOne = "Apache POI - Component Overview";
	
	public final String paragOneListOne = "The Apache POI project is the master project for developing pure Java ports of file formats based on Microsoft's OLE 2"+
			" Compound Document Format. OLE 2 Compound Document Format is used by Microsoft Office Documents, as well as by programs using MFC"+
			" property sets to serialize their document objects.";
	
	public final String paragTwoListOne = "Apache POI is also the master project for developing pure Java ports of file formats based on Office Open XML (ooxml)."+
			" OOXML is part of an ECMA / ISO standardisation effort. This documentation is quite large, but you can normally find the bit you need"+
			" without too much effort! ECMA-376 standard is here, and is also under the Microsoft OSP.";
	
	public final String paragTheeListOne = "POIFS is the oldest and most stable part of POI. It is our port of the OLE 2 Compound Document Format to pure Java."+
			" It supports both read and write functionality. All of our components for the binary (non-XML) Microsoft Office formats ultimately"+
			" rely on it by definition. Please see the POIFS project page for more information.";
	
	//List Two
	
	public final String headerListTwo = "List Numbering Overview";
	public final String paragOneListTwo = "Adding a second list Numbering to Word Document";
	public final String paragTwoListTwo = "The numbering is working correctly";
	public final String paragThreeListTwo = "We just addeda second list Numbering to the Word Document";
	
	
	public List<String> paragraphsSectionOne (){
		
		List<String> list = List.of(headerListOne, paragOneListOne, paragTwoListOne, paragTheeListOne);
		
		return list;
	}
	
	public List<String> paragraphsSectionTwo (){
		
		List<String> list = List.of(headerListTwo, paragOneListTwo, paragTwoListTwo, paragThreeListTwo);
		
		return list;
	}
	

}
