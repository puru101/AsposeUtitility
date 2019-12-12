package org.sae.authoring;

import java.io.File;

import org.jsoup.Jsoup;
import org.sae.authoring.api.constants.Constants;
import org.sae.authoring.api.util.AsposeUtils;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.SaveFormat;

public class AsposeQuery {
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			AsposeUtils.checkLicence();
			convertHTMLToDocX_MultipleSections(Constants.TEMP_HTML_FILE_1);
			System.out.println("done");


		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void convertHTMLToDocX_MultipleSections(String docName) {
		// Loading document from specified document name and path.

		try {

			String templatePath = Constants.DOCUMENT_TEMPLATE + "category-3.docx";
			com.aspose.words.Document dstDoc = new com.aspose.words.Document(templatePath);

			com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(dstDoc);
			builder.moveToDocumentEnd();			
			// Need to add table of contents 
		    ParagraphFormat paragraphFormat = builder.getParagraphFormat();

			paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
			builder.writeln();
			builder.writeln("TABLE OF CONTENTS");
			builder.writeln();
			builder.insertTableOfContents(" \\o \"1-3\" \\h \\z \\u ");			
			
			builder.insertBreak(BreakType.PAGE_BREAK);			
			
			
			// Below for loop for getting HTML content from database
			for (int i = 1; i <= 3; i++) {				
				File input = new File(Constants.SOURCE_DIR + docName + i+".html");
				org.jsoup.nodes.Document doc1 = Jsoup.parse(input, "UTF-8", "http://example.com/");
				builder.insertHtml(doc1.html(), true);	
			}
			dstDoc.updateFields();
			dstDoc.save(Constants.SOURCE_DIR + "output.docx", SaveFormat.DOCX);			

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	
	
  
	public static void convertHTMLToDocX(String docName) {
		// Loading document from specified document name and path.
		Document doc;

		try {
			
			doc = new Document(Constants.SOURCE_DIR + docName);
			
			
			//File input = new File(Constants.SOURCE_DIR + docName);			
			
			//doc = AsposeUtils.convertXMLTODocument(readFileToByteArray(input));
			
			
			//org.jsoup.nodes.Document doc1 = Jsoup.parse(input, "UTF-8", "http://example.com/");

			String templatePath = Constants.DOCUMENT_TEMPLATE + "category-1.docx";
			com.aspose.words.Document dstDoc = new com.aspose.words.Document(templatePath);
			System.out.println("default font:"+dstDoc.getStyles().getDefaultFont().getName());
			
			com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(dstDoc);

			builder.moveToDocumentEnd();
			ParagraphFormat paragraphFormat = builder.getParagraphFormat();

			paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
			builder.writeln();
			builder.writeln("TABLE OF CONTENTS");
			builder.writeln();
			builder.insertTableOfContents(" \\o \"1-3\" \\h \\z \\u ");
			// Start the actual document content on the second page.
			builder.insertBreak(BreakType.PAGE_BREAK);

			//com.aspose.words.Document srcDoc = AsposeUtils
				//	.convertXMLTODocument(AsposeUtils.convertHtmlToXml(doc1.html()));
			//builder.appendDocument(doc, true);
			dstDoc.appendDocument(doc, ImportFormatMode.USE_DESTINATION_STYLES);
			dstDoc.updateFields();
			dstDoc.joinRunsWithSameFormatting();
			
			dstDoc.copyStylesFromTemplate(templatePath);
			dstDoc.save(Constants.SOURCE_DIR + "output_2.docx", SaveFormat.DOCX);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
