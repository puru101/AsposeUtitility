package org.sae.authoring.api.util;

import co.elastic.apm.api.CaptureSpan;
import com.aspose.words.*;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.sae.authoring.api.constants.Constants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.slf4j.Marker;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;
import java.util.*;

public class AsposeUtils {

	private static final Logger logger = LoggerFactory.getLogger(AsposeUtils.class);
	private static Marker marker;
	private static final String DIR = "D:/2019/5/138/82d20329-b7ff-4f13-a5a3-55d25895b108/";

	/**
	 * This method is used to validate Aspose API license
	 * 
	 * @throws Exception
	 */
	public static void checkLicence() throws Exception {
		License license = new License();
		try {
			license.setLicense(Constants.LICENSE_DIR);	
			logger.trace("Aspose License is Set!");
		
		} catch (Exception e) {
			logger.error("Error in Aspose Licence check {}", e);
			throw e;
		}
	}

	/**
	 * This method is used to convert HTML data to XML format which will use to
	 * generate ms-word document
	 * 
	 * @param htmlData,
	 *            is input for method Code Reference:
	 *            https://apireference.aspose.com/java/words/com.aspose.words/Node
	 * @return
	 */
	@CaptureSpan
	public static ByteArrayOutputStream convertHTMLToXML(String htmlData) throws Exception {
		AsposeUtils.checkLicence();
		logger.trace("convertHTMLToXML Start {}");
		ByteArrayOutputStream xmlOutputStream = new ByteArrayOutputStream();
		try {
			// Creating document.
			Document doc = new Document();

			// inserting html string to document using Document Builder.
			new DocumentBuilder(doc).insertHtml(htmlData);

			// Save the document to stream
			doc.save(xmlOutputStream, SaveFormat.FLAT_OPC);
		
		} catch (Exception e) {
			logger.trace("Error form convertHTMLToXML {}",e);
		}
		logger.trace("convertHTMLToXML end {}");
		return xmlOutputStream;
	}

	/**
	 * This method converts html into xml
	 * 
	 * @param htmlString
	 * @return it returns XmlOutputStream as Byte[].
	 * @throws Exception
	 */
	public static byte[] convertHtmlToXml(String htmlString) throws Exception {

		logger.trace("Inside convert html to xml method");

		// Creating document.
		Document doc = new Document();

		// inserting html string to document using Document Builder.
		new DocumentBuilder(doc).insertHtml(htmlString);

		// Create a new memory stream.
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		// Save the document to stream
		doc.save(outputStream, SaveFormat.FLAT_OPC);

		// Convert the document to byte form.
		byte[] xmlOutputStream = outputStream.toByteArray();
		// The bytes are now ready to be stored/transmitted.

		// return outputStream byte array;
		logger.trace("end convertHtmlToXml {}");
		return xmlOutputStream;
	}

	/**
	 * This method converts the document into xml.
	 * 
	 * @param docName
	 * @return it returns OutputStream byte array.
	 * @throws Exception
	 */
	public static byte[] convertDocToXml(String docName) throws Exception {
		logger.trace("Inside convert doc to xml method");

		// Load the document.
		Document doc = new Document(Constants.SOURCE_DIR + docName);

		// Create a new memory stream.
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		// Save the document to stream
		doc.save(outputStream, SaveFormat.FLAT_OPC);

		// Convert the document to byte form.
		byte[] xmlOutputStream = outputStream.toByteArray();
		// The bytes are now ready to be stored/transmitted.

		// return outputStream;
		return xmlOutputStream;
	}

	/**
	 * This method is used to convert HTML data to XML format which will use to
	 * generate ms-word document
	 * 
	 * @param htmlData,
	 *            is input for method Code Reference:
	 *            https://apireference.aspose.com/java/words/com.aspose.words/Node
	 * @return
	 */
	public static Document convertXMLTODocument(byte[] xmlData) {

		try {
			ByteArrayInputStream docInStream = null;
			// if (xmlData != null) {
			docInStream = new ByteArrayInputStream(xmlData);
			// } else {
			// User will able to download blank document by adding raw data
			// docInStream = new
			// ByteArrayInputStream(AsposeUtils.convertHTMLToXML("<p>Default
			// section version data</p>").toByteArray());
			// }

			// Create document from stream
			Document outDoc = new Document(docInStream);

			logger.trace("Document string data {}" + outDoc);
			return outDoc;
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * This method converts the document into html.
	 * 
	 * @param docName
	 * @return it returns html string.
	 * @throws Exception
	 */
	public static String convertDocToHtml(String docName) throws Exception {

		logger.trace("Inside convert doc to html method");

		// Loading document from specified document name and path.
		Document doc = new Document(Constants.SOURCE_DIR + docName);

		// To extract section with numbering
		// Create an instance of HtmlSaveOptions and set a few options.
		HtmlSaveOptions saveOptions = new HtmlSaveOptions();
		saveOptions.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
		saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.PER_SECTION);
		saveOptions.setExportRelativeFontSize(true);

		// When this property is set to true image data is
		// exported directly on the img elements and separate files are not
		// created.
		saveOptions.setExportImagesAsBase64(true);

		// Create a new memory stream.
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		// Save the document to stream
		doc.save(outputStream, SaveFormat.FLAT_OPC);

		// Convert the document to byte form.
		byte[] xmlOutputStream = outputStream.toByteArray();
		// The bytes are now ready to be stored/transmitted.

		// Now reverse the steps to load the bytes back into a document object.
		ByteArrayInputStream inputStream = new ByteArrayInputStream(xmlOutputStream);

		// Load the stream into a new document object.
		Document loadDoc = new Document(inputStream);

		// Converting document as single node.
		Node node = loadDoc;

		// Convert the document to HTML and return as a string.
		// Pass the instance of
		// HtmlSaveOptions to
		// to use the specified options during the conversion.
		return node.toString(saveOptions);
	}

	/**
	 * This method converts the xml into html.
	 * 
	 * @param docBytes
	 * @return it returns html string.
	 * @throws Exception
	 */
	public static String convertXmlToHtml(byte[] docBytes) throws Exception {

		logger.trace("Inside convert xml to html method");

		// Load the bytes back into a document object.
		ByteArrayInputStream inputStream = new ByteArrayInputStream(docBytes);

		// Load the stream into a new document object.
		Document doc = new Document(inputStream);

		// To extract section with numbering
		// Create an instance of HtmlSaveOptions and set a few options.
		HtmlSaveOptions saveOptions = new HtmlSaveOptions();
		saveOptions.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
		saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.PER_SECTION);
		saveOptions.setExportRelativeFontSize(true);

		// When this property is set to true image data is
		// exported directly on the img elements and separate files are not
		// created.
		saveOptions.setExportImagesAsBase64(true);

		// Converting document as single node.
		Node node = doc;

		// Convert the document to HTML and return as a string.
		// Pass the instance of
		// HtmlSaveOptions to
		// to use the specified options during the conversion.
		return node.toString(saveOptions);
	}

	public static void convertDocToHTML(String docName) {
		logger.trace("Inside convert doc to html method");

		// Loading document from specified document name and path.
		Document doc;
		try {
			doc = new Document("c:\\" + Constants.SOURCE_DIR + docName);

			com.aspose.words.HtmlSaveOptions options = new com.aspose.words.HtmlSaveOptions();
			options.setSaveFormat(com.aspose.words.SaveFormat.HTML);
			options.setExportImagesAsBase64(true);
			options.setExportFontsAsBase64(true);
			ByteArrayOutputStream baos1 = new ByteArrayOutputStream();
			doc.getRange().replace(ControlChar.NON_BREAKING_SPACE, "", new FindReplaceOptions());
			doc.save("c:\\" + Constants.SOURCE_DIR+"abc.html" , options);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	@Deprecated
	public static DocumentSectionBean extractSections_old(UUID docId, InputStream inputStream, UUID fileUploadId,
			List<org.sae.authoring.data.model.Section> authorSections,
			List<org.sae.authoring.api.model.Section> allDocSections) throws Exception {

		try {

			logger.trace("In Extract Sections Method inputStreamResourceForDoc: " + inputStream);

			// checks license for Aspose library
			AsposeUtils.checkLicence();

			DocumentSectionBean documentSectionBean = new DocumentSectionBean();

			Map<org.sae.authoring.data.model.Section, SectionData> sectionDataMap = new HashMap<>();

			List<String> uploadedSections = new ArrayList<>();

			Map<String, org.sae.authoring.data.model.Section> foundSections = new HashMap<>();

			logger.trace("Before getting document");
			// Gets uploaded document object

			Document doc = null;
			if (inputStream != null) {

				try {
					// InputStream inputStream = new
					// FileInputStream(inputStreamResourceForDoc.getFile());
					logger.trace("In Extract Sections Method inputStream " + inputStream);
					doc = new Document(inputStream);
				} catch (Exception e) {
					logger.error("Error while ACCESSING DOCUMENT" + e);
				}
			}

			logger.trace("Gets uploaded document object {}", doc);
			doc.updateListLabels();
			DocumentBuilder builder = new DocumentBuilder(doc);

			int bmo = 1;
			for (Section section : doc.getSections()) {
				if (doc.getFirstSection() == section)
					continue;

				builder.moveTo(((Section) section.getPreviousSibling()).getBody().getLastParagraph());
				int orientation = section.getPageSetup().getOrientation();

				// When we extract sections (e.g. TEST), extractContent method
				// does
				// not retain
				// Section Brakes in extracted
				// sections, i.e.The extractContent method does not extract the
				// section breaks.
				// We are adding bookmarks to know the position of section
				// brakes.
				if (section.getPageSetup().getSectionStart() == SectionStart.CONTINUOUS) {
					builder.startBookmark("BM_BreakC" + bmo);
					builder.endBookmark("BM_BreakC" + bmo);
					builder.startBookmark(orientation + "Orientation" + bmo);
					builder.endBookmark(orientation + "Orientation" + bmo);
				}
				if (section.getPageSetup().getSectionStart() == SectionStart.NEW_PAGE) {
					builder.startBookmark("BM_BreakNewPage" + bmo);
					builder.endBookmark("BM_BreakNewPage" + bmo);
					builder.startBookmark(orientation + "Orientation" + bmo);
					builder.endBookmark(orientation + "Orientation" + bmo);
				}
				bmo++;
			}

			int i = 1;

			NodeCollection<Paragraph> nodes = doc.getChildNodes(NodeType.PARAGRAPH, true);

			for (Paragraph para : (Iterable<Paragraph>) nodes) {

				if (para.getParagraphFormat().isHeading()
						&& para.getParagraphFormat().getStyle().getName().equals(Constants.HEADING_STYLE)) {

					logger.trace(marker, "Section {} has header style {} ", para.getText(),
							para.getParagraphFormat().getStyle().getName());

					uploadedSections.add(para.getText().toLowerCase().trim());

					Paragraph paragraph = new Paragraph(doc);

					para.getParentNode().insertBefore(paragraph, para);

					builder.moveTo(paragraph);

					builder.startBookmark(Constants.BOOKMARK_NAME + i);

					builder.endBookmark(Constants.BOOKMARK_NAME + i);

					for (org.sae.authoring.data.model.Section sec : authorSections) {

						if (para.getText().toLowerCase().trim().equals(sec.getTitle().toLowerCase().trim())) {

							foundSections.put(Constants.BOOKMARK_NAME + i, sec);

							logger.trace(marker, "Sections Found while extracting :{} {}", sec.getTitle(),
									para.getText());

						}

					}
					// increase counter
					i++;

				}

			}

			logger.trace(marker, "uploadedSections {}", uploadedSections);
			logger.trace(marker, "foundSections {}", foundSections.values().toArray());

			builder.moveToDocumentEnd();
			builder.startBookmark(Constants.BOOKMARK_NAME + i);
			builder.endBookmark(Constants.BOOKMARK_NAME + i);

			// Creating xml and html files separately because when we extract
			// xml we
			// are not
			// extracting section with numbering but when we are extracting html
			// we
			// are
			// extracting section with numbering
			//generateXml(fileUploadId, sectionDataMap, foundSections, doc, i, docId);
			generateHtml(fileUploadId, sectionDataMap, foundSections, doc, i, docId);

			List<String> unmatchedSections = unmatchedSections(allDocSections, uploadedSections);
			logger.trace(marker, "unmatchedSections {}. ", unmatchedSections);

			List<org.sae.authoring.data.model.Section> notFoundSections = notFoundSections(authorSections,
					uploadedSections);
			logger.trace(marker, "notFoundSections {} ", notFoundSections);

			logger.trace(marker, "SectionDataMap {} ", sectionDataMap);

			documentSectionBean.setNotFoundSections(notFoundSections);
			documentSectionBean.setUnmatchedSections(unmatchedSections);
			documentSectionBean.setSectionData(sectionDataMap);

			return documentSectionBean;

		} catch (Exception e) {
			logger.error("ERROR IN EXTRACTING DOCUMENT", e);
			throw e;
		}
	}

	public static DocumentSectionBean extractSections(UUID docId, InputStream inputStream, UUID fileUploadId,
			org.sae.authoring.data.model.Section section) throws Exception {

		logger.trace("Start AsposeUtils.extractSections() inputStreamResourceForDoc{}", inputStream);

		// checks license for Aspose library
		AsposeUtils.checkLicence();

		DocumentSectionBean documentSectionBean = new DocumentSectionBean();

		Map<org.sae.authoring.data.model.Section, SectionData> sectionDataMap = new HashMap<>();

		List<String> uploadedSections = new ArrayList<>();

		Map<String, org.sae.authoring.data.model.Section> foundSections = new HashMap<>();

		logger.trace("AsposeUtils.extractSections() Before getting document");
		// Gets uploaded document object

		Document doc = null;
		if (inputStream != null) {
			doc = new Document(inputStream);
		}
		
		logger.trace("AsposeUtils.extractSections() Remove Nonbreaking Space Characters");
		
		
		

		logger.trace("AsposeUtils.extractSections() uploaded document object {}", doc);
		doc.updateListLabels();
		
		//TODO: for ckeditor
		doc.joinRunsWithSameFormatting();
		
		DocumentBuilder builder = new DocumentBuilder(doc);	
		
		
		int bmo = 1;
		for (Section asposeSection : doc.getSections()) {
			if (doc.getFirstSection() == asposeSection)
				continue;

			builder.moveTo(((Section) asposeSection.getPreviousSibling()).getBody().getLastParagraph());
			int orientation = asposeSection.getPageSetup().getOrientation();

			// When we extract sections (e.g. TEST), extractContent method
			// does
			// not retain
			// Section Brakes in extracted
			// sections, i.e.The extractContent method does not extract the
			// section breaks.
			// We are adding bookmarks to know the position of section
			// brakes.
			if (asposeSection.getPageSetup().getSectionStart() == SectionStart.CONTINUOUS) {
				builder.startBookmark("BM_BreakC" + bmo);
				builder.endBookmark("BM_BreakC" + bmo);
				builder.startBookmark(orientation + "Orientation" + bmo);
				builder.endBookmark(orientation + "Orientation" + bmo);
			}
			if (asposeSection.getPageSetup().getSectionStart() == SectionStart.NEW_PAGE) {
				builder.startBookmark("BM_BreakNewPage" + bmo);
				builder.endBookmark("BM_BreakNewPage" + bmo);
				builder.startBookmark(orientation + "Orientation" + bmo);
				builder.endBookmark(orientation + "Orientation" + bmo);
			}
			bmo++;
		}

		int i = 1;

		NodeCollection<Paragraph> nodes = doc.getChildNodes(NodeType.PARAGRAPH, true);

		for (Paragraph para : (Iterable<Paragraph>) nodes) {

			if (para.getParagraphFormat().isHeading()
					&& para.getParagraphFormat().getStyle().getName().equals(Constants.HEADING_STYLE)) {

				logger.trace("AsposeUtils.extractSections() {}",marker, "Section {} has header style {} ", para.getText(),
						para.getParagraphFormat().getStyle().getName());

				uploadedSections.add(para.getText().toLowerCase().trim());

				Paragraph paragraph = new Paragraph(doc);

				para.getParentNode().insertBefore(paragraph, para);

				builder.moveTo(paragraph);

				builder.startBookmark(Constants.BOOKMARK_NAME + i);

				builder.endBookmark(Constants.BOOKMARK_NAME + i);

				if (para.getText().toLowerCase().trim().equals(section.getTitle().toLowerCase().trim())) {

					foundSections.put(Constants.BOOKMARK_NAME + i, section);

					logger.trace("AsposeUtils.extractSections() {}",marker, "Sections Found while extracting : {}", section.getTitle(), para.getText());

				}

				// increase counter
				i++;

			}

		}

		logger.trace(marker, "uploadedSections {}", uploadedSections);
		logger.trace("AsposeUtils.extractSections() {}",marker, "foundSections {}", foundSections.values().toArray());

		builder.moveToDocumentEnd();
		builder.startBookmark(Constants.BOOKMARK_NAME + i);
		builder.endBookmark(Constants.BOOKMARK_NAME + i);

		// Creating xml and html files separately because when we extract
		// xml we
		// are not
		// extracting section with numbering but when we are extracting html
		// we
		// are
		// extracting section with numbering
		
		
				
		generateXml(fileUploadId, sectionDataMap, foundSections, doc, i, docId);
		generateHtml(fileUploadId, sectionDataMap, foundSections, doc, i, docId);

		/*
		 * List<String> unmatchedSections = unmatchedSections(allDocSections,
		 * uploadedSections); logger.trace(marker, "unmatchedSections {}. ",
		 * unmatchedSections);
		 * 
		 * List<org.sae.authoring.data.model.Section> notFoundSections =
		 * notFoundSections(authorSections, uploadedSections);
		 */
		// logger.trace(marker, "notFoundSections {} ", notFoundSections);

		logger.trace(marker, "SectionDataMap {} ", sectionDataMap);

		// documentSectionBean.setNotFoundSections(notFoundSections);
		// documentSectionBean.setUnmatchedSections(unmatchedSections);
		documentSectionBean.setSectionData(sectionDataMap);
		logger.trace("End AsposeUtils.extractSections() documentSectionBean {}", documentSectionBean);
		return documentSectionBean;

	}

	private static void generateXml(UUID fileUploadId,
			Map<org.sae.authoring.data.model.Section, SectionData> sectionDataMap,
			Map<String, org.sae.authoring.data.model.Section> foundSections, Document doc, int i, UUID docId)
			throws Exception {

		for (int bm = 1; bm < i; bm++) {
			BookmarkStart bookmarkStart = doc.getRange().getBookmarks().get(Constants.BOOKMARK_NAME + bm)
					.getBookmarkStart();

			BookmarkStart bookmarkEnd = doc.getRange().getBookmarks().get(Constants.BOOKMARK_NAME + (bm + 1))
					.getBookmarkStart();

			// First extract the content between these nodes including the
			// bookmark.

			ArrayList<Node> extractedNodes = extractContent(bookmarkStart, bookmarkEnd, false);

			Document dstDoc = generateDocument(doc, extractedNodes);

			dstDoc.getFirstSection().getPageSetup().setOrientation(
					((Section) bookmarkStart.getAncestor(NodeType.SECTION)).getPageSetup().getOrientation());
			dstDoc.getFirstSection().getPageSetup().setLeftMargin(
					((Section) bookmarkStart.getAncestor(NodeType.SECTION)).getPageSetup().getLeftMargin());
			dstDoc.getFirstSection().getPageSetup().setRightMargin(
					((Section) bookmarkStart.getAncestor(NodeType.SECTION)).getPageSetup().getRightMargin());
			dstDoc.getFirstSection().getPageSetup().setTopMargin(
					((Section) bookmarkStart.getAncestor(NodeType.SECTION)).getPageSetup().getTopMargin());
			dstDoc.getFirstSection().getPageSetup().setBottomMargin(
					((Section) bookmarkStart.getAncestor(NodeType.SECTION)).getPageSetup().getBottomMargin());

			// We are inserting the section brakes as in source docx.
			DocumentBuilder dstbuilder = new DocumentBuilder(dstDoc);
			for (Bookmark bookmark : dstDoc.getRange().getBookmarks()) {
				if (bookmark.getName().contains("BM_BreakC")) {
					dstbuilder.moveTo(bookmark.getBookmarkEnd().getNextSibling());
					dstbuilder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
				} else if (bookmark.getName().contains("BM_BreakNewPage")) {
					dstbuilder.moveTo(bookmark.getBookmarkEnd());
					dstbuilder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
				} else if (bookmark.getName().contains("Orientation")) {
					if (bookmark.getName().startsWith("1"))
						((Section) bookmark.getBookmarkStart().getAncestor(NodeType.SECTION)).getPageSetup()
								.setOrientation(Orientation.PORTRAIT);
					else
						((Section) bookmark.getBookmarkStart().getAncestor(NodeType.SECTION)).getPageSetup()
								.setOrientation(Orientation.LANDSCAPE);
				}
			}

			for (Bookmark bookmark : dstDoc.getRange().getBookmarks()) {
				if (bookmark.getName().contains("BM_BreakC")) {
					bookmark.remove();
				} else if (bookmark.getName().contains("BM_BreakNewPage")) {
					bookmark.remove();
				} else if (bookmark.getName().contains("Orientation")) {
					bookmark.getBookmarkStart().getParentNode().remove();
				}
			}

			String bkName = Constants.BOOKMARK_NAME + bm;

			foundSections.forEach((k, v) -> {
				try {
					if (bkName.equals(k)) {

						SectionData sectionData = new SectionData();

						ByteArrayOutputStream xmlOutputStream = new ByteArrayOutputStream();
						
						//TODO: convert 
						/*FontSettings fontSettings = new FontSettings();
						fontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
						dstDoc.setFontSettings(fontSettings);*/
					
							
						dstDoc.save(xmlOutputStream, SaveFormat.FLAT_OPC);

						logger.trace("generateXml() setting xmlOutputStream to sectionData");

						sectionData.setSectionOutputStream(xmlOutputStream);

						sectionDataMap.put(v, sectionData);

					}
				} catch (Exception e) {
					logger.error(marker, "Exception in converting XML to ByteArrayOutputStream" + e);
					throw new RuntimeException(e);

				}

			});

		}
	}

	private static void generateHtml(UUID fileUploadId,
			Map<org.sae.authoring.data.model.Section, SectionData> sectionDataMap,
			Map<String, org.sae.authoring.data.model.Section> foundSections, Document doc, int i, UUID docId)
			throws Exception {
		
		logger.trace("AsposeUtils.generateHtml() fileUploadId {}",fileUploadId);
		
		//TODO: commented for testing
		// To extract section with numbering
		//HtmlSaveOptions options = new HtmlSaveOptions();
		//options.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
		//options.setExportImagesAsBase64(true);
		
		
		//doc.save(Constants.DEST_DIR + "docInHtml" + Constants.XML_EXTENSION, options);

		//doc = new Document(Constants.DEST_DIR + "docInHtml" + Constants.XML_EXTENSION);
		//TODO: commented for testing
		
		
		for (int bm = 1; bm < i; bm++) {
			BookmarkStart bookmarkStart = doc.getRange().getBookmarks().get(Constants.BOOKMARK_NAME + bm)
					.getBookmarkStart();

			BookmarkStart bookmarkEnd = doc.getRange().getBookmarks().get(Constants.BOOKMARK_NAME + (bm + 1))
					.getBookmarkStart();

			// First extract the content between these nodes including the
			// bookmark.

			ArrayList<Node> extractedNodes = extractContent(bookmarkStart, bookmarkEnd, false);

			Document dstDoc = generateDocument(doc, extractedNodes);

			String bkName = Constants.BOOKMARK_NAME + bm;

			foundSections.forEach((k, v) -> {
				try {
					if (k.equals(bkName)) {

						logger.trace("In foundSections [{},{}] ", bkName, k);

						//Remove Nonbreaking Space Characters
						//TODO: for ckeditor
						dstDoc.joinRunsWithSameFormatting();
						dstDoc.getRange().replace(ControlChar.NON_BREAKING_SPACE, "", new FindReplaceOptions());
						
						Node node = dstDoc;

						// Create an instance of HtmlSaveOptions and set a few
						// options.
						HtmlSaveOptions saveOptions = new HtmlSaveOptions();
						saveOptions.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
						saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE);
						saveOptions.setExportRelativeFontSize(true);

						// When this property is set to true image data is
						// exported directly on the img
						// elements and separate files are not created.
						saveOptions.setExportImagesAsBase64(true);

						//TODO:for ckeditor, minimize inline styloe
						saveOptions.setPrettyFormat(true);
						saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
						
						
						// Convert the document to HTML and return as a string.
						// Pass the instance of
						// HtmlSaveOptions to
						// to use the specified options during the conversion.
						String nodeAsHtml = node.toString(saveOptions);

						
						// remove all inline-style from html content
						org.jsoup.nodes.Document htmlDocument = Jsoup.parse(nodeAsHtml);
						Elements els = htmlDocument.getAllElements();
						for (Element e : els) {
							e.removeAttr("style");
							//e.removeAttr("img");
						}
						logger.trace("AsposeUtils.generateHtml() remove all inline style htmlDocument {}",htmlDocument);
						
						
						dstDoc.save(Constants.DEST_DIR + v.getTitle() + Constants.HTML_EXTENSION, saveOptions);

						sectionDataMap.forEach((sec, secData) -> {

							if (v.getId() == sec.getId()) {

								secData.setSectionHtmlString(htmlDocument.html());
								logger.trace("AsposeUtils.generateHtml() remove all inline style {}",secData.getSectionHtmlString());
							}

						});

					}
				} catch (Exception e) {
					logger.error(marker, "Exception in converting HTML to string {}", e);
					throw new RuntimeException(e);

				}

			});

		}
	}

	/**
	 * 
	 * This method retrieves sections that are present in author section list
	 * but not in uploaded doc section list.
	 * 
	 * Assigned Sections that are not uploaded by Author
	 * 
	 * @param authorSections
	 * @param uploadedSections
	 */
	private static List<org.sae.authoring.data.model.Section> notFoundSections(
			List<org.sae.authoring.data.model.Section> authorSections, List<String> uploadedSections) throws Exception {

		List<org.sae.authoring.data.model.Section> notFoundSections = new ArrayList<>();

		logger.debug(marker, "uploadedSection {}, authorSectionTitles {}", authorSections, uploadedSections);

		for (org.sae.authoring.data.model.Section authorSection : authorSections) {
			if (!uploadedSections.contains(authorSection.getTitle().toLowerCase().trim())) {
				notFoundSections.add(authorSection);
			}
		}
		return notFoundSections;

	}

	/**
	 * This Method retrieves sections that are present in uploaded document but
	 * not in all Doc sections list.
	 * 
	 * Author adds new Section without sponsor's approval
	 * 
	 * @param allDocSections
	 * @param uploadedSections
	 * @return
	 */
	private static List<String> unmatchedSections(List<org.sae.authoring.api.model.Section> allDocSections,
			List<String> uploadedSections) throws Exception {
		List<String> unmatchedSections = new ArrayList<>();
		List<String> allDocSectionsTitles = new ArrayList<>();

		allDocSections.forEach(s -> allDocSectionsTitles.add(s.getTitle().toLowerCase().trim()));

		logger.debug("uploadedSection {}", uploadedSections);
		logger.debug("allDocSections {}", allDocSections);

		for (String uploadedSection : uploadedSections) {
			if (!allDocSectionsTitles.contains(uploadedSection)) {
				unmatchedSections.add(uploadedSection);
			}
		}

		return unmatchedSections;
	}

	/**
	 * Extracts a range of nodes from a document found between specified markers
	 * and returns a copy of those nodes. Content can be extracted between
	 * inline nodes, block level nodes, and also special nodes such as Comment
	 * or Boomarks. Any combination of different marker types can used.
	 *
	 * @param startNode
	 *            The node which defines where to start the extraction from the
	 *            document. This node can be block or inline level of a body.
	 * @param endNode
	 *            The node which defines where to stop the extraction from the
	 *            document. This node can be block or inline level of body.
	 * @param isInclusive
	 *            Should the marker nodes be included.
	 */
	public static ArrayList extractContent(Node startNode, Node endNode, boolean isInclusive) throws Exception {
		// First check that the nodes passed to this method are valid for use.
		verifyParameterNodes(startNode, endNode);
		CompositeNode cloneNode = null;

		// Create a list to store the extracted nodes.
		ArrayList nodes = new ArrayList();

		// Keep a record of the original nodes passed to this method so we can
		// split marker nodes if needed.
		Node originalStartNode = startNode;
		Node originalEndNode = endNode;

		// Extract content based on block level nodes (paragraphs and tables).
		// Traverse through parent nodes to find them.
		// We will split the content of first and last nodes depending if the
		// marker nodes are inline
		while (startNode.getParentNode().getNodeType() != NodeType.BODY)
			startNode = startNode.getParentNode();

		while (endNode.getParentNode().getNodeType() != NodeType.BODY)
			endNode = endNode.getParentNode();

		boolean isExtracting = true;
		boolean isStartingNode = true;
		boolean isEndingNode;
		// The current node we are extracting from the document.
		Node currNode = startNode;

		// Begin extracting content. Process all block level nodes and
		// specifically split the first and last nodes when needed so paragraph
		// formatting is retained.
		// Method is little more complex than a regular extractor as we need to
		// factor in extracting using inline nodes, fields, bookmarks etc as to
		// make it really useful.
		while (isExtracting) {

			try {
				// Clone the current node and its children to obtain a copy.
				cloneNode = (CompositeNode) currNode.deepClone(true);
			} catch (Exception e) {
				logger.error("THERE IS SECTION HEADING WITHOUT NAME [{}{}]", currNode.getText(), e);
			}

			isEndingNode = currNode.equals(endNode);

			if (isStartingNode || isEndingNode) {
				// We need to process each marker separately so pass it off to a
				// separate method instead.
				if (isStartingNode) {
					processMarker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode);
					isStartingNode = false;
				}

				// Conditional needs to be separate as the block level start and
				// end markers maybe the same node.
				if (isEndingNode) {
					processMarker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode);
					isExtracting = false;
				}
			} else
				// Node is not a start or end marker, simply add the copy to the
				// list.
				nodes.add(cloneNode);

			// Move to the next node and extract it. If next node is null that
			// means the rest of the content is found in a different section.
			if (currNode.getNextSibling() == null && isExtracting) {
				// Move to the next section.
				Section nextSection = (Section) currNode.getAncestor(NodeType.SECTION).getNextSibling();
				currNode = nextSection.getBody().getFirstChild();
			} else {
				// Move to the next node in the body.
				currNode = currNode.getNextSibling();
			}
		}

		// Return the nodes between the node markers.
		return nodes;
	}

	public static Document generateDocument(Document srcDoc, ArrayList nodes) throws Exception {
		// Create a blank document.
		Document dstDoc = new Document();
		// Remove the first paragraph from the empty document.
		dstDoc.getFirstSection().getBody().removeAllChildren();

		// Import each node from the list into the new document. Keep the
		// original formatting of the node.
		NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

		for (Node node : (Iterable<Node>) nodes) {
			Node importNode = importer.importNode(node, true);
			dstDoc.getFirstSection().getBody().appendChild(importNode);
		}

		// Return the generated document.
		return dstDoc;
	}

	private static void verifyParameterNodes(Node startNode, Node endNode) throws Exception {
		// The order in which these checks are done is important.
		if (startNode == null)
			throw new IllegalArgumentException("Start node cannot be null");
		if (endNode == null)
			throw new IllegalArgumentException("End node cannot be null");

		if (!startNode.getDocument().equals(endNode.getDocument()))
			throw new IllegalArgumentException("Start node and end node must belong to the same document");

		if (startNode.getAncestor(NodeType.BODY) == null || endNode.getAncestor(NodeType.BODY) == null)
			throw new IllegalArgumentException("Start node and end node must be a child or descendant of a body");

		// Check the end node is after the start node in the DOM tree
		// First check if they are in different sections, then if they're not
		// check their position in the body of the same section they are in.
		Section startSection = (Section) startNode.getAncestor(NodeType.SECTION);
		Section endSection = (Section) endNode.getAncestor(NodeType.SECTION);

		int startIndex = startSection.getParentNode().indexOf(startSection);
		int endIndex = endSection.getParentNode().indexOf(endSection);

		if (startIndex == endIndex) {
			if (startSection.getBody().indexOf(startNode) > endSection.getBody().indexOf(endNode))
				throw new IllegalArgumentException("The end node must be after the start node in the body");
		} else if (startIndex > endIndex)
			throw new IllegalArgumentException("The section of end node must be after the section start node");
	}

	/**
	 * Removes the content before or after the marker in the cloned node
	 * depending on the type of marker.
	 */
	private static void processMarker(CompositeNode cloneNode, ArrayList nodes, Node node, boolean isInclusive,
			boolean isStartMarker, boolean isEndMarker) throws Exception {
		// If we are dealing with a block level node just see if it should be
		// included and add it to the list.
		if (!isInline(node)) {
			// Don't add the node twice if the markers are the same node
			if (!(isStartMarker && isEndMarker)) {
				if (isInclusive)
					nodes.add(cloneNode);
			}
			return;
		}

		// If a marker is a FieldStart node check if it's to be included or not.
		// We assume for simplicity that the FieldStart and FieldEnd appear in
		// the same paragraph.
		if (node.getNodeType() == NodeType.FIELD_START) {
			// If the marker is a start node and is not be included then skip to
			// the end of the field.
			// If the marker is an end node and it is to be included then move
			// to the end field so the field will not be removed.
			if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive)) {
				while (node.getNextSibling() != null && node.getNodeType() != NodeType.FIELD_END)
					node = node.getNextSibling();

			}
		}

		// If either marker is part of a comment then to include the comment
		// itself we need to move the pointer forward to the Comment
		// node found after the CommentRangeEnd node.
		if (node.getNodeType() == NodeType.COMMENT_RANGE_END) {
			while (node.getNextSibling() != null && node.getNodeType() != NodeType.COMMENT)
				node = node.getNextSibling();

		}

		// Find the corresponding node in our cloned node by index and return
		// it.
		// If the start and end node are the same some child nodes might already
		// have been removed. Subtract the
		// difference to get the right index.
		int indexDiff = node.getParentNode().getChildNodes().getCount() - cloneNode.getChildNodes().getCount();

		// Child node count identical.
		if (indexDiff == 0)
			node = cloneNode.getChildNodes().get(node.getParentNode().indexOf(node));
		else
			node = cloneNode.getChildNodes().get(node.getParentNode().indexOf(node) - indexDiff);

		// Remove the nodes up to/from the marker.
		boolean isSkip;
		boolean isProcessing = true;
		boolean isRemoving = isStartMarker;
		Node nextNode = cloneNode.getFirstChild();

		while (isProcessing && nextNode != null) {
			Node currentNode = nextNode;
			isSkip = false;

			if (currentNode.equals(node)) {
				if (isStartMarker) {
					isProcessing = false;
					if (isInclusive)
						isRemoving = false;
				} else {
					isRemoving = true;
					if (isInclusive)
						isSkip = true;
				}
			}

			nextNode = nextNode.getNextSibling();
			if (isRemoving && !isSkip)
				currentNode.remove();
		}

		// After processing the composite node may become empty. If it has don't
		// include it.
		if (!(isStartMarker && isEndMarker)) {
			if (cloneNode.hasChildNodes())
				nodes.add(cloneNode);
		}

	}

	/**
	 * Checks if a node passed is an inline node.
	 */
	private static boolean isInline(Node node) throws Exception {
		// Test if the node is descendant of a Paragraph or Table node and also
		// is not a paragraph or a table a paragraph inside a comment class
		// which is descendant of a pararaph is possible.
		return ((node.getAncestor(NodeType.PARAGRAPH) != null || node.getAncestor(NodeType.TABLE) != null)
				&& !(node.getNodeType() == NodeType.PARAGRAPH || node.getNodeType() == NodeType.TABLE));
	}
}
