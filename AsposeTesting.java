package org.sae.authoring;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.OffsetDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.sae.authoring.api.constants.Constants;
import org.sae.authoring.api.util.AsposeUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.aspose.words.Body;
import com.aspose.words.BreakType;
import com.aspose.words.ControlChar;
import com.aspose.words.CssStyleSheetType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ExportHeadersFootersMode;
import com.aspose.words.ExportListLabels;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.HtmlOfficeMathOutputMode;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.SaveFormat;
import com.aspose.words.Style;

public class AsposeTesting {

	private static final Logger logger = LoggerFactory.getLogger(AsposeTesting.class);

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			AsposeUtils.checkLicence();
			// convertDocToHTMLNBSP("category-1.docx");
			convertDocToHTMLUsingJSOUP("MS Template_9.26.2016.docx");
			//convertHTMLToDocX("Titanium Alloy, Welding.doc.html");

			
			// convertHtmlToXml("abc.html");
			System.out.println("done");
			// Document doc = new Document(Constants.DEST_DIR + "docInHtml" +
			// Constants.XML_EXTENSION);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static String getPreviousDateData() {
		String format="EOM";
	     int count=2;
		
		
		// here you can pass your own date
		/*Date myDate = null;
		String input = "2009-09-30";
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
		try {
			myDate = dateFormat.parse(input);
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}*/
		
		Date myDate = new Date();	
		Calendar cal1 = Calendar.getInstance();
		cal1.setTime(myDate);
		

		if (format.equals("EOD")) {
			cal1.add(Calendar.DAY_OF_YEAR, -count);
		}
		if (format.equals("EOM")) {
			cal1.add(Calendar.MONTH, -count);
		}
		Date previousDate = cal1.getTime();
		System.out.println("previousDate :"+previousDate);
		return previousDate.toString();

	}

	public static void checkMathML(String docName) {

		// Loading document from specified document name and path.
		Document doc = null;
		try {
			doc = new Document(Constants.SOURCE_DIR + docName);
			doc.joinRunsWithSameFormatting();

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void convertDocToHTMLUsingJSOUP(String docName) {

		// Loading document from specified document name and path.
		Document doc;

		try {
			doc = new Document(Constants.SOURCE_DIR + docName);
			doc.joinRunsWithSameFormatting();
			Document dstDoc = new Document();
			doc.getRange().replace(ControlChar.NON_BREAKING_SPACE, "", new FindReplaceOptions());
			// Create an instance of HtmlSaveOptions and set a few
			// options.
			HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
			saveOptions.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
			saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE);
			saveOptions.setExportRelativeFontSize(true);
			
			//Adding mathml code on generated html.
			saveOptions.setOfficeMathOutputMode(HtmlOfficeMathOutputMode.MATH_ML);

			// When this property is set to true image data is
			// exported directly on the img
			// elements and separate files are not created.
			saveOptions.setExportImagesAsBase64(true);
			saveOptions.setPrettyFormat(true);
			
			// TODO: minimize number of inline styloe
			saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
			doc.save(Constants.SOURCE_DIR + Constants.TEMP_HTML_FILE, saveOptions);

			File input = new File(Constants.SOURCE_DIR + Constants.TEMP_HTML_FILE);

			Path path = Paths.get(input.getAbsolutePath());

			// File input = new File(Constants.SOURCE_DIR + "Aerospace
			// Template.html");
			org.jsoup.nodes.Document doc1 = Jsoup.parse(input, "UTF-8");

			// Elements elements = doc1.getAllElements();

			//Elements elements = doc1.select("span[style*=font-style:italic]");
			//elements = doc1.select("p[style*=margin-bottom:12pt]");
			// Elements elements = doc1.select("span[style*=font-weight:bold;
			// font-style:italic]");

			// for (Element e : elements) {
			// elements.tagName("strong");
			// elements.tagName("i");
			// e.removeAttr("style");
			// Element pre=e;
			// System.out.println("pre outer : " + e);

			// }

			// elements = doc1.select("span[style*=-aw-import:ignore]");
			// for (Element e : elements) {
			// elements.tagName("strong");
			// elements.tagName("i");
			// e.removeAttr("style");
			// Element pre=e;
			// System.out.println("pre outer : " + e);

			// }
			//
			// elements =
			// doc1.select("span[style*=text-decoration:underlined]");
			//
			// for (Element e : elements) {
			// elements.tagName("u");
			// // Element pre=e;
			// System.out.println("pre outer : " + e);
			//
			// }
			//
			// elements = doc1.select("span[style*=font-weight:bold]");
			//
			// for (Element e : elements) {
			// elements.tagName("strong");
			// // Element pre=e;
			// System.out.println("pre outer : " + e);
			//
			// }
			// elements = doc1.select("span[style*=font-style:italic]");
			// for (Element e : elements) {
			// elements.tagName("i");
			// // Element pre=e;
			// System.out.println("pre outer : " + e);
			//
			// }
			//
			// elements =
			// doc1.select("span[style*=text-decoration:line-through]");
			// for (Element e : elements) {
			// elements.tagName("s");
			// // Element pre=e;
			// System.out.println("pre outer : .*[^0-9].*" + e);
			//
			// }
			//
			// Elements elements = doc1.getAllElements();
			// for (Element e : elements) {
			// e.removeAttr("style");
			// }

			Elements elementHeadings = doc1.select("h1");
			Pattern pattern = Pattern.compile("\\d.*");

			for (Element e : elementHeadings) {
				// e.child(0).remove();
			}

			 Elements subHeadings = doc1.select("h1,h2,h3,h4,h5,h6");
			 for (Element e : subHeadings) {
	
			 Elements spanTags = e.select("span");
			 for (Element span : spanTags) {
				 System.out.println(span.html());
				 System.out.println(pattern.matcher(span.text()).matches());
			 if (pattern.matcher(span.text()).matches()) {
				// System.out.println(span.html());
				 e.child(0).remove();
			 }
			 if (span.text().equalsIgnoreCase(Constants.NON_BREAKING_SPACE)) {
			 e.child(0).remove();
			 }
			
			 }
			 }

			 
			 elementHeadings = doc1.select("span[style*=font-family:Symbol]");
				
				for (Element e : elementHeadings) {					
					//e.remove();			
					System.out.println("pre outer : " + e);

				}
				
				
			 Elements  paragraps = doc1.select("p");
			 
			 pattern = Pattern.compile("AHead1|AHead2|AHead3|AHead4|AHead5|AHead6");
			 for (Element e : paragraps) {
			 
			 if (pattern.matcher(e.className()).matches()) {
				 System.out.println(e.className());
				 e.child(0).remove();
			
			 }
			 }
			 
			 
			//dstDoc.save(Constants.DEST_DIR + docName+ Constants.HTML_EXTENSION, saveOptions);
			
			final File f = new File(Constants.SOURCE_DIR + docName +Constants.HTML_EXTENSION);
			FileUtils.writeStringToFile(f, doc1.html(), "UTF-8");
			 Files.delete(path);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static byte[] readFileToByteArray(File file) {
		FileInputStream fis = null;
		// Creating a byte array using the length of the file
		// file.length returns long which is cast to int
		byte[] bArray = new byte[(int) file.length()];
		try {
			fis = new FileInputStream(file);
			fis.read(bArray);
			fis.close();

		} catch (IOException ioExp) {
			ioExp.printStackTrace();
		}
		return bArray;
	}

	public static void convertHTMLToDocX(String docName) {
		// Loading document from specified document name and path.
		Document doc;

		try {

			doc = new Document(Constants.SOURCE_DIR + docName);
			File input = new File(Constants.SOURCE_DIR + docName);

			org.jsoup.nodes.Document doc1 = Jsoup.parse(input, "UTF-8", "http://example.com/");

			Element divConntent = doc1.select("body div").first();

			String templatePath = Constants.DOCUMENT_TEMPLATE + "category-3.docx";
			com.aspose.words.Document dstDoc = new com.aspose.words.Document(templatePath);

			com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(dstDoc);
			builder.moveToDocumentEnd();
			ParagraphFormat paragraphFormat = builder.getParagraphFormat();

			paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
			builder.writeln();
			builder.writeln("TABLE OF CONTENTS");
			builder.writeln();
			builder.insertTableOfContents(" \\o \"1-3\" \\h \\z \\u ");
			// dstDoc.updateFields();
			// builder.writeln();
			// builder.insertTableOfContents(" \\h \\z \\t \"Heading 7,8,Heading
			// 8,9,Appendix,7\" ");

			// Start the actual document content on the second page.
			builder.insertBreak(BreakType.PAGE_BREAK);
			builder.getParagraphFormat().clearFormatting();
			// System.out.println("Xml
			// output:"+AsposeUtils.convertHtmlToXml(doc1.html()).toString());
			com.aspose.words.Document srcDoc = AsposeUtils
					.convertXMLTODocument(AsposeUtils.convertHtmlToXml(divConntent.html()));

			builder.insertHtml(divConntent.html(), true);
			// System.out.println("Xml output:"+srcDoc.getText());
			// dstDoc.appendDocument(srcDoc,
			// ImportFormatMode.USE_DESTINATION_STYLES);
			// builder.insertDocument(srcDoc,
			// ImportFormatMode.USE_DESTINATION_STYLES);

			// dstDoc.copyStylesFromTemplate(templatePath);
			// dstDoc.joinRunsWithSameFormatting();

			dstDoc.updateFields();
			dstDoc.save(Constants.SOURCE_DIR + "expected result.docx", SaveFormat.DOCX);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static String readAllText(String filePath) {
		String content = "";
		try {
			content = new String(Files.readAllBytes(Paths.get(filePath)));
		} catch (IOException e) {
			e.printStackTrace();
		}

		return content;
	}

	public static void convertHTMLToDocX_Tahir(String docName) {
		// Loading document from specified document name and path.
		Document doc;

		try {

			doc = new Document(Constants.SOURCE_DIR + docName);
			File input = new File(Constants.SOURCE_DIR + docName);

			// doc =
			// AsposeUtils.convertXMLTODocument(readFileToByteArray(input));

			org.jsoup.nodes.Document doc1 = Jsoup.parse(input, "UTF-8", "http://example.com/");

			String templatePath = Constants.DOCUMENT_TEMPLATE + "category-3.docx";
			com.aspose.words.Document dstDoc = new com.aspose.words.Document(templatePath);

			DocumentBuilder builder = new DocumentBuilder(dstDoc);
			builder.moveToDocumentEnd();
			builder.insertHtml(readAllText(Constants.SOURCE_DIR + docName), true);

			doc.save(Constants.SOURCE_DIR + "output.docx");

			//
			// com.aspose.words.DocumentBuilder builder = new
			// com.aspose.words.DocumentBuilder(dstDoc);
			//
			// builder.moveToDocumentEnd();
			// ParagraphFormat paragraphFormat = builder.getParagraphFormat();
			//
			// paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
			// builder.writeln();
			// builder.writeln("TABLE OF CONTENTS");
			// builder.writeln();
			// builder.insertTableOfContents(" \\o \"1-3\" \\h \\z \\u ");
			// // Start the actual document content on the second page.
			// builder.insertBreak(BreakType.PAGE_BREAK);
			//
			// com.aspose.words.Document srcDoc = AsposeUtils
			// .convertXMLTODocument(AsposeUtils.convertHtmlToXml(doc1.html()));
			// dstDoc.appendDocument(srcDoc,
			// ImportFormatMode.USE_DESTINATION_STYLES);
			// dstDoc.updateFields();
			// dstDoc.joinRunsWithSameFormatting();
			//
			// dstDoc.copyStylesFromTemplate(templatePath);
			// dstDoc.save(Constants.SOURCE_DIR + "output_2.docx",
			// SaveFormat.DOCX);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void convertHTMLToDocX_12102019(String docName) {
		// Loading document from specified document name and path.
		Document doc;

		try {

			doc = new Document(Constants.SOURCE_DIR + docName);
			File input = new File(Constants.SOURCE_DIR + docName);

			// doc =
			// AsposeUtils.convertXMLTODocument(readFileToByteArray(input));

			org.jsoup.nodes.Document doc1 = Jsoup.parse(input, "UTF-8", "http://example.com/");

			String templatePath = Constants.DOCUMENT_TEMPLATE + "category-3.docx";
			com.aspose.words.Document dstDoc = new com.aspose.words.Document(templatePath);
			System.out.println("default font:" + dstDoc.getStyles().getDefaultFont().getName());

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

			com.aspose.words.Document srcDoc = AsposeUtils
					.convertXMLTODocument(AsposeUtils.convertHtmlToXml(doc1.html()));

			NodeCollection nodes = srcDoc.getChildNodes(NodeType.BODY, true);
			for (Body node : (Iterable<Body>) nodes) {

				for (Paragraph paragraph : node.getParagraphs()) {
					if (paragraph.getParagraphFormat().getStyle().getName().equalsIgnoreCase("Heading 1")
							|| paragraph.getParagraphFormat().getStyle().getName().equalsIgnoreCase("Heading 2")) {
						System.out.println("default font:" + paragraph.getText());
						Style style = paragraph.getParagraphFormat().getStyle();
						style.getFont().setSize(10);
						paragraph.getParagraphFormat().setStyle(style);
					}

				}
			}

			dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);
			dstDoc.copyStylesFromTemplate(templatePath);
			dstDoc.joinRunsWithSameFormatting();
			dstDoc.updateFields();

			dstDoc.save(Constants.SOURCE_DIR + "output_2.docx", SaveFormat.DOCX);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void convertDocToHTML_OLD_10102019(String docName) {

		// Loading document from specified document name and path.
		Document doc;

		try {
			doc = new Document(Constants.SOURCE_DIR + docName);

			// NodeCollection nodes = doc.getChildNodes(NodeType.BODY, true);
			// Look through all paragraphs to find those with the specified
			// style.

			// Font font = builder.getFont();
			// font.setName("Arial");

			/*
			 * for (Body node : (Iterable<Body>) nodes) {
			 * 
			 * for (Paragraph paragraph : node.getParagraphs()) {
			 * System.out.println("style {}"+paragraph.getParagraphFormat().
			 * getStyle().getName()); if
			 * (paragraph.getParagraphFormat().getStyle().getName().equals(
			 * "Body")){
			 * 
			 * 
			 * paragraph.getParagraphFormat().getStyle().getFont().setBold(false
			 * );
			 * paragraph.getParagraphFormat().getStyle().getFont().setSize(10);
			 * paragraph.getParagraphFormat().getStyle().getFont().setName(
			 * "Arial"); //paragraph.getParagraphFormat().
			 * builder.getParagraphFormat().setStyle(paragraph.
			 * getParagraphFormat().getStyle());
			 * 
			 * }else if (paragraph.getParagraphFormat().getStyle().getName().
			 * equalsIgnoreCase("Heading 1")){
			 * 
			 * 
			 * paragraph.getParagraphFormat().getStyle().getFont().setBold(false
			 * );
			 * paragraph.getParagraphFormat().getStyle().getFont().setSize(20);
			 * paragraph.getParagraphFormat().getStyle().getFont().setName(
			 * "Arial"); builder.getParagraphFormat().setStyle(paragraph.
			 * getParagraphFormat().getStyle()); }else if
			 * (paragraph.getParagraphFormat().getStyle().getName().
			 * equalsIgnoreCase("Heading 2")){
			 * 
			 * paragraph.getParagraphFormat().getStyle().getFont().setBold(false
			 * );
			 * paragraph.getParagraphFormat().getStyle().getFont().setSize(15);
			 * paragraph.getParagraphFormat().getStyle().getFont().setName(
			 * "Arial"); builder.getParagraphFormat().setStyle(paragraph.
			 * getParagraphFormat().getStyle()); }else if
			 * (paragraph.getParagraphFormat().getStyle().getName().
			 * equalsIgnoreCase("DocList")){
			 * 
			 * 
			 * paragraph.getParagraphFormat().getStyle().getFont().setBold(false
			 * ); paragraph.getParagraphFormat().getStyle().getFont().setItalic(
			 * true);
			 * paragraph.getParagraphFormat().getStyle().getFont().setSize(11);
			 * paragraph.getParagraphFormat().getStyle().getFont().setName(
			 * "Arial"); builder.getParagraphFormat().setStyle(paragraph.
			 * getParagraphFormat().getStyle()); }else if
			 * (paragraph.getParagraphFormat().getStyle().getName().
			 * equalsIgnoreCase("Normal")){
			 * 
			 * paragraph.getParagraphFormat().getStyle().getFont().setBold(false
			 * );
			 * paragraph.getParagraphFormat().getStyle().getFont().setSize(10);
			 * paragraph.getParagraphFormat().getStyle().getFont().setName(
			 * "Arial"); builder.getParagraphFormat().setStyle(paragraph.
			 * getParagraphFormat().getStyle()); }else{
			 * 
			 * paragraph.getParagraphFormat().getStyle().getFont().setBold(false
			 * );
			 * paragraph.getParagraphFormat().getStyle().getFont().setSize(10);
			 * paragraph.getParagraphFormat().getStyle().getFont().setName(
			 * "Arial"); builder.getParagraphFormat().setStyle(paragraph.
			 * getParagraphFormat().getStyle()); } } }
			 */

			// doc.joinRunsWithSameFormatting();
			// HtmlSaveOptions options = new HtmlSaveOptions();
			// options.setPrettyFormat(true);
			// options.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);

			File input = new File(Constants.SOURCE_DIR + docName);
			org.jsoup.nodes.Document doc1 = Jsoup.parse(input, "UTF-8", "http://example.com/");

			// org.jsoup.nodes.Document doc1 = Jsoup.connect();

			com.aspose.words.Document dstDoc = new com.aspose.words.Document(Constants.SOURCE_DIR + "category-1.docx");

			com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(dstDoc);

			builder.moveToDocumentEnd();
			ParagraphFormat paragraphFormat = builder.getParagraphFormat();

			paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
			builder.writeln();
			builder.writeln("TABLE OF CONTENTS");
			builder.writeln();
			builder.insertTableOfContents(" \\o \"1-3\" \\h \\z \\u ");

			// builder.writeln();
			// builder.insertTableOfContents(" \\h \\z \\t \"Heading 7,8,Heading
			// 8,9,Appendix,7\" ");

			// Start the actual document content on the second page.
			builder.insertBreak(BreakType.PAGE_BREAK);

			// doc.save(MyDir + "19.9.html", options);

			// com.aspose.words.HtmlSaveOptions options = new
			// com.aspose.words.HtmlSaveOptions();
			// options.setSaveFormat(com.aspose.words.SaveFormat.HTML);
			// options.setExportImagesAsBase64(true);

			// ByteArrayOutputStream baos1 = new ByteArrayOutputStream();
			// FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
			// doc.getRange().replace("TEST testing data", "code for testing
			// text data", findReplaceOptions);

			com.aspose.words.Document srcDoc = AsposeUtils
					.convertXMLTODocument(AsposeUtils.convertHtmlToXml(doc1.html()));
			dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);
			dstDoc.updateFields();
			// dstDoc.joinRunsWithSameFormatting();
			String templatePath = Constants.SOURCE_DIR + "category-1.docx";
			dstDoc.copyStylesFromTemplate(templatePath);
			// doc.save( Constants.SOURCE_DIR +
			// "output_1.docx",options);

			dstDoc.save(Constants.SOURCE_DIR + "output_1.docx", SaveFormat.DOCX);

			// dstDoc.save( Constants.SOURCE_DIR + "output_1.docx");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void convertHtmlToXml(String docName) throws Exception {

		logger.trace("Inside convert html to xml method");

		// Creating document.
		Document doc = new Document(Constants.SOURCE_DIR + docName);

		// inserting html string to document using Document Builder.
		com.aspose.words.DocumentBuilder builder = new com.aspose.words.DocumentBuilder(doc);
		NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);

		// Look through all paragraphs to find those with the specified style.
		for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs) {
			if (paragraph.getParagraphFormat().getStyle().getName().equals("Body")) {
				System.out.println("done {}" + builder.getDocument().getStyles().get(0).getBaseStyleName());
				// Style paraStyle =
				// builder.getDocument().getStyles().get("Heading 1");
				paragraph.getParagraphFormat().getStyle().getFont().setBold(false);
				paragraph.getParagraphFormat().getStyle().getFont().setSize(14);
				paragraph.getParagraphFormat().getStyle().getFont().setName("Times New Roman");

				builder.getParagraphFormat().setStyle(paragraph.getParagraphFormat().getStyle());

			} else {
				Style paraStyle1 = builder.getDocument().getStyles().get("Heading 1");

				System.out.println("done else {}" + builder.getDocument().getStyles().get(0).getBaseStyleName());

				paraStyle1.getFont().setBold(false);
				paraStyle1.getFont().setSize(20);
				paraStyle1.getFont().setName("Times New Roman");
				builder.getParagraphFormat().setStyleName(paraStyle1.getName());
			}

		}
		// Create a new memory stream.
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		// Save the document to stream
		doc.save(outputStream, SaveFormat.FLAT_OPC);

		doc.save(Constants.SOURCE_DIR + "abc.docx");
	}

	public static void convertDocToHTMLNBSP(String docName) {

		// Loading document from specified document name and path.
		Document doc;
		try {
			doc = new Document(Constants.SOURCE_DIR + docName);
			// com.aspose.words.DocumentBuilder builder = new
			// com.aspose.words.DocumentBuilder(doc);
			// doc.getRange().replace("&nbsp", "", new
			// FindReplaceOptions(FindReplaceDirection.FORWARD));

			com.aspose.words.HtmlSaveOptions options = new com.aspose.words.HtmlSaveOptions();
			options.setSaveFormat(com.aspose.words.SaveFormat.HTML);
			options.setExportImagesAsBase64(true);

			// ByteArrayOutputStream baos1 = new ByteArrayOutputStream();
			// FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
			// doc.getRange().replace("TEST testing data", "code for testing
			// text data", findReplaceOptions);

			// doc.getRange().replace(ControlChar.NON_BREAKING_SPACE, "", new
			// FindReplaceOptions());
			doc.save(Constants.SOURCE_DIR + "abc.html", options);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	/*
	 * 
	 * private static void
	 * generateHtml(Map<org.sae.authoring.data.model.Section, SectionData>
	 * sectionDataMap, Map<String, org.sae.authoring.data.model.Section>
	 * foundSections, Document doc, int i) throws Exception {
	 * 
	 * 
	 * // To extract section with numbering HtmlSaveOptions options = new
	 * HtmlSaveOptions();
	 * options.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
	 * options.setExportImagesAsBase64(true);
	 * 
	 * doc.save(Constants.DEST_DIR + "docInHtml" + Constants.XML_EXTENSION,
	 * options);
	 * 
	 * doc = new Document(Constants.DEST_DIR + "docInHtml" +
	 * Constants.XML_EXTENSION);
	 * 
	 * for (int bm = 1; bm < i; bm++) { BookmarkStart bookmarkStart =
	 * doc.getRange().getBookmarks().get(Constants.BOOKMARK_NAME + bm)
	 * .getBookmarkStart();
	 * 
	 * BookmarkStart bookmarkEnd =
	 * doc.getRange().getBookmarks().get(Constants.BOOKMARK_NAME + (bm + 1))
	 * .getBookmarkStart();
	 * 
	 * // First extract the content between these nodes including the //
	 * bookmark.
	 * 
	 * ArrayList<Node> extractedNodes = extractContent(bookmarkStart,
	 * bookmarkEnd, false);
	 * 
	 * Document dstDoc = generateDocument(doc, extractedNodes);
	 * 
	 * String bkName = Constants.BOOKMARK_NAME + bm;
	 * 
	 * foundSections.forEach((k, v) -> { try { if (k.equals(bkName)) {
	 * 
	 * logger.trace("In foundSections [{},{}] ", bkName, k);
	 * 
	 * //Remove Nonbreaking Space Characters //
	 * dstDoc.getRange().replace(ControlChar.NON_BREAKING_SPACE, "", new
	 * FindReplaceOptions());
	 * 
	 * Node node = dstDoc;
	 * node.getRange().replace(ControlChar.NON_BREAKING_SPACE, "", new
	 * FindReplaceOptions()); // Create an instance of HtmlSaveOptions and set a
	 * few // options. HtmlSaveOptions saveOptions = new HtmlSaveOptions();
	 * saveOptions.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
	 * saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.
	 * PER_SECTION); saveOptions.setExportRelativeFontSize(true);
	 * 
	 * // When this property is set to true image data is // exported directly
	 * on the img // elements and separate files are not created.
	 * saveOptions.setExportImagesAsBase64(true);
	 * 
	 * // Convert the document to HTML and return as a string. // Pass the
	 * instance of // HtmlSaveOptions to // to use the specified options during
	 * the conversion. String nodeAsHtml = node.toString(saveOptions);
	 * 
	 * 
	 * dstDoc.save(Constants.DEST_DIR + v.getTitle() + Constants.HTML_EXTENSION,
	 * saveOptions);
	 * 
	 * sectionDataMap.forEach((sec, secData) -> {
	 * 
	 * if (v.getId() == sec.getId()) {
	 * 
	 * secData.setSectionHtmlString(nodeAsHtml);
	 * 
	 * }
	 * 
	 * });
	 * 
	 * } } catch (Exception e) { //logger.error(marker,
	 * "Exception in converting HTML to string {}", e); throw new
	 * RuntimeException(e);
	 * 
	 * }
	 * 
	 * });
	 * 
	 * } } private static void verifyParameterNodes(Node startNode, Node
	 * endNode) throws Exception { // The order in which these checks are done
	 * is important. if (startNode == null) throw new
	 * IllegalArgumentException("Start node cannot be null"); if (endNode ==
	 * null) throw new IllegalArgumentException("End node cannot be null");
	 * 
	 * if (!startNode.getDocument().equals(endNode.getDocument())) throw new
	 * IllegalArgumentException("Start node and end node must belong to the same document"
	 * );
	 * 
	 * if (startNode.getAncestor(NodeType.BODY) == null ||
	 * endNode.getAncestor(NodeType.BODY) == null) throw new
	 * IllegalArgumentException("Start node and end node must be a child or descendant of a body"
	 * );
	 * 
	 * // Check the end node is after the start node in the DOM tree // First
	 * check if they are in different sections, then if they're not // check
	 * their position in the body of the same section they are in. Section
	 * startSection = (Section) startNode.getAncestor(NodeType.SECTION); Section
	 * endSection = (Section) endNode.getAncestor(NodeType.SECTION);
	 * 
	 * int startIndex = startSection.getParentNode().indexOf(startSection); int
	 * endIndex = endSection.getParentNode().indexOf(endSection);
	 * 
	 * if (startIndex == endIndex) { if
	 * (startSection.getBody().indexOf(startNode) >
	 * endSection.getBody().indexOf(endNode)) throw new
	 * IllegalArgumentException("The end node must be after the start node in the body"
	 * ); } else if (startIndex > endIndex) throw new
	 * IllegalArgumentException("The section of end node must be after the section start node"
	 * ); } public static ArrayList extractContent(Node startNode, Node endNode,
	 * boolean isInclusive) throws Exception { // First check that the nodes
	 * passed to this method are valid for use. verifyParameterNodes(startNode,
	 * endNode); CompositeNode cloneNode = null;
	 * 
	 * // Create a list to store the extracted nodes. ArrayList nodes = new
	 * ArrayList();
	 * 
	 * // Keep a record of the original nodes passed to this method so we can //
	 * split marker nodes if needed. Node originalStartNode = startNode; Node
	 * originalEndNode = endNode;
	 * 
	 * // Extract content based on block level nodes (paragraphs and tables). //
	 * Traverse through parent nodes to find them. // We will split the content
	 * of first and last nodes depending if the // marker nodes are inline while
	 * (startNode.getParentNode().getNodeType() != NodeType.BODY) startNode =
	 * startNode.getParentNode();
	 * 
	 * while (endNode.getParentNode().getNodeType() != NodeType.BODY) endNode =
	 * endNode.getParentNode();
	 * 
	 * boolean isExtracting = true; boolean isStartingNode = true; boolean
	 * isEndingNode; // The current node we are extracting from the document.
	 * Node currNode = startNode;
	 * 
	 * // Begin extracting content. Process all block level nodes and //
	 * specifically split the first and last nodes when needed so paragraph //
	 * formatting is retained. // Method is little more complex than a regular
	 * extractor as we need to // factor in extracting using inline nodes,
	 * fields, bookmarks etc as to // make it really useful. while
	 * (isExtracting) {
	 * 
	 * try { // Clone the current node and its children to obtain a copy.
	 * cloneNode = (CompositeNode) currNode.deepClone(true); } catch (Exception
	 * e) { logger.error("THERE IS SECTION HEADING WITHOUT NAME [{}{}]",
	 * currNode.getText(), e); }
	 * 
	 * isEndingNode = currNode.equals(endNode);
	 * 
	 * if (isStartingNode || isEndingNode) { // We need to process each marker
	 * separately so pass it off to a // separate method instead. if
	 * (isStartingNode) { processMarker(cloneNode, nodes, originalStartNode,
	 * isInclusive, isStartingNode, isEndingNode); isStartingNode = false; }
	 * 
	 * // Conditional needs to be separate as the block level start and // end
	 * markers maybe the same node. if (isEndingNode) { processMarker(cloneNode,
	 * nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode);
	 * isExtracting = false; } } else // Node is not a start or end marker,
	 * simply add the copy to the // list. nodes.add(cloneNode);
	 * 
	 * // Move to the next node and extract it. If next node is null that //
	 * means the rest of the content is found in a different section. if
	 * (currNode.getNextSibling() == null && isExtracting) { // Move to the next
	 * section. Section nextSection = (Section)
	 * currNode.getAncestor(NodeType.SECTION).getNextSibling(); currNode =
	 * nextSection.getBody().getFirstChild(); } else { // Move to the next node
	 * in the body. currNode = currNode.getNextSibling(); } }
	 * 
	 * // Return the nodes between the node markers. return nodes; }
	 * 
	 * public static Document generateDocument(Document srcDoc, ArrayList nodes)
	 * throws Exception { // Create a blank document. Document dstDoc = new
	 * Document(); // Remove the first paragraph from the empty document.
	 * dstDoc.getFirstSection().getBody().removeAllChildren();
	 * 
	 * // Import each node from the list into the new document. Keep the //
	 * original formatting of the node. NodeImporter importer = new
	 * NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
	 * 
	 * for (Node node : (Iterable<Node>) nodes) { Node importNode =
	 * importer.importNode(node, true);
	 * dstDoc.getFirstSection().getBody().appendChild(importNode); }
	 * 
	 * // Return the generated document. return dstDoc; } private static void
	 * processMarker(CompositeNode cloneNode, ArrayList nodes, Node node,
	 * boolean isInclusive, boolean isStartMarker, boolean isEndMarker) throws
	 * Exception { // If we are dealing with a block level node just see if it
	 * should be // included and add it to the list. if (!isInline(node)) { //
	 * Don't add the node twice if the markers are the same node if
	 * (!(isStartMarker && isEndMarker)) { if (isInclusive)
	 * nodes.add(cloneNode); } return; }
	 * 
	 * // If a marker is a FieldStart node check if it's to be included or not.
	 * // We assume for simplicity that the FieldStart and FieldEnd appear in //
	 * the same paragraph. if (node.getNodeType() == NodeType.FIELD_START) { //
	 * If the marker is a start node and is not be included then skip to // the
	 * end of the field. // If the marker is an end node and it is to be
	 * included then move // to the end field so the field will not be removed.
	 * if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive)) {
	 * while (node.getNextSibling() != null && node.getNodeType() !=
	 * NodeType.FIELD_END) node = node.getNextSibling();
	 * 
	 * } }
	 * 
	 * // If either marker is part of a comment then to include the comment //
	 * itself we need to move the pointer forward to the Comment // node found
	 * after the CommentRangeEnd node. if (node.getNodeType() ==
	 * NodeType.COMMENT_RANGE_END) { while (node.getNextSibling() != null &&
	 * node.getNodeType() != NodeType.COMMENT) node = node.getNextSibling();
	 * 
	 * }
	 * 
	 * // Find the corresponding node in our cloned node by index and return //
	 * it. // If the start and end node are the same some child nodes might
	 * already // have been removed. Subtract the // difference to get the right
	 * index. int indexDiff = node.getParentNode().getChildNodes().getCount() -
	 * cloneNode.getChildNodes().getCount();
	 * 
	 * // Child node count identical. if (indexDiff == 0) node =
	 * cloneNode.getChildNodes().get(node.getParentNode().indexOf(node)); else
	 * node = cloneNode.getChildNodes().get(node.getParentNode().indexOf(node) -
	 * indexDiff);
	 * 
	 * // Remove the nodes up to/from the marker. boolean isSkip; boolean
	 * isProcessing = true; boolean isRemoving = isStartMarker; Node nextNode =
	 * cloneNode.getFirstChild();
	 * 
	 * while (isProcessing && nextNode != null) { Node currentNode = nextNode;
	 * isSkip = false;
	 * 
	 * if (currentNode.equals(node)) { if (isStartMarker) { isProcessing =
	 * false; if (isInclusive) isRemoving = false; } else { isRemoving = true;
	 * if (isInclusive) isSkip = true; } }
	 * 
	 * nextNode = nextNode.getNextSibling(); if (isRemoving && !isSkip)
	 * currentNode.remove(); }
	 * 
	 * // After processing the composite node may become empty. If it has don't
	 * // include it. if (!(isStartMarker && isEndMarker)) { if
	 * (cloneNode.hasChildNodes()) nodes.add(cloneNode); }
	 * 
	 * }
	 *//**
		 * Checks if a node passed is an inline node.
		 *//*
		 * private static boolean isInline(Node node) throws Exception { // Test
		 * if the node is descendant of a Paragraph or Table node and also // is
		 * not a paragraph or a table a paragraph inside a comment class //
		 * which is descendant of a pararaph is possible. return
		 * ((node.getAncestor(NodeType.PARAGRAPH) != null ||
		 * node.getAncestor(NodeType.TABLE) != null) && !(node.getNodeType() ==
		 * NodeType.PARAGRAPH || node.getNodeType() == NodeType.TABLE)); }
		 */
}
