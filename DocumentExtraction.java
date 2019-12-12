package org.sae.authoring;

import java.util.ArrayList;
import java.util.List;

import org.sae.authoring.api.constants.Constants;
import org.sae.authoring.api.util.AsposeUtils;
import org.slf4j.Marker;

import com.aspose.words.Bookmark;
import com.aspose.words.BookmarkStart;
import com.aspose.words.BreakType;
import com.aspose.words.CompositeNode;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.License;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeImporter;
import com.aspose.words.NodeType;
import com.aspose.words.Orientation;
import com.aspose.words.Paragraph;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
import com.aspose.words.SectionStart;

public class DocumentExtraction {
	private static Marker marker;

	public static void main(String[] args) {
	
		try {
			AsposeUtils.checkLicence();
			extractSectionsForPartTemplate("AS28938A.doc");			
			System.out.println("done");

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	public static void extractSectionsForPartTemplate(String docName) throws Exception {
		Document doc = new Document("C:\\sw-files\\source\\" + docName);
		 java.util.List<String> uploadedSections = new ArrayList<>();

		System.out.print("AsposeUtils.extractSections() Before getting document");
		// Gets uploaded document object

		System.out.print("AsposeUtils.extractSections() Remove Nonbreaking Space Characters");

		System.out.print("AsposeUtils.extractSections() uploaded document object {}"+ doc);
		doc.updateListLabels();

		// TODO: for ckeditor
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
		 
		for(Section section : doc.getSections())
		{
		    NodeCollection<Paragraph> nodes = section.getBody().getChildNodes(NodeType.PARAGRAPH, true);

		    for (Paragraph para : (Iterable<Paragraph>) nodes) {

		        if (para.getParagraphFormat().isHeading()
		                && para.getParagraphFormat().getStyle().getName().equals(Constants.HEADING_STYLE)) {

		            uploadedSections.add(para.getText().toLowerCase().trim());

		            Paragraph paragraph = new Paragraph(doc);

		            para.getParentNode().insertBefore(paragraph, para);

		            builder.moveTo(paragraph);

		            builder.startBookmark(Constants.BOOKMARK_NAME + i);

		            builder.endBookmark(Constants.BOOKMARK_NAME + i);

		            // increase counter
		            i++;

		        }

		    }

		    System.out.print("uploadedSections {}"+ uploadedSections);
		    builder.moveToDocumentEnd();
		    builder.startBookmark(Constants.BOOKMARK_NAME + i);
		    builder.endBookmark(Constants.BOOKMARK_NAME + i);
			Document dstDoc = new Document();
			
			//DocumentBuilder dstbuilder = new DocumentBuilder(dstDoc);
			doc.save("C:\\sw-files\\source\\test.doc" );
			
		    System.out.print("AsposeUtils.extractSections() going for extract html content and xml {}");
		    generateXml(doc, i);
		}
	}

	public static void extractSections(String docName) throws Exception {

		System.out.print("Start AsposeUtils.extractSections() inputStreamResourceForDoc{}"+ docName);

		Document doc = new Document("C:\\sw-files\\source\\" + docName);
			
		// checks license for Aspose library
		AsposeUtils.checkLicence();

		List<String> uploadedSections = new ArrayList<>();

		System.out.print("AsposeUtils.extractSections() Before getting document");
		// Gets uploaded document object

		System.out.print("AsposeUtils.extractSections() Remove Nonbreaking Space Characters");

		System.out.print("AsposeUtils.extractSections() uploaded document object {}"+ doc);
		doc.updateListLabels();

		// TODO: for ckeditor
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

				uploadedSections.add(para.getText().toLowerCase().trim());

				Paragraph paragraph = new Paragraph(doc);

				para.getParentNode().insertBefore(paragraph, para);

				builder.moveTo(paragraph);

				builder.startBookmark(Constants.BOOKMARK_NAME + i);

				builder.endBookmark(Constants.BOOKMARK_NAME + i);

				// increase counter
				i++;

			}

		}

		System.out.print("uploadedSections {}"+ uploadedSections);

		builder.moveToDocumentEnd();
		builder.startBookmark(Constants.BOOKMARK_NAME + i);
		builder.endBookmark(Constants.BOOKMARK_NAME + i);


		System.out.print("AsposeUtils.extractSections() going for extract html content and xml {}");

		
		Document dstDoc = new Document();
		
		//DocumentBuilder dstbuilder = new DocumentBuilder(dstDoc);
		doc.save("C:\\sw-files\\source\\test.doc" );
		
		generateXml(doc, i);


	}

	private static void generateXml(Document doc, int i) throws Exception {

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



			//ByteArrayOutputStream xmlOutputStream = new ByteArrayOutputStream();

		
			dstDoc.save("C:\\sw-files\\source\\test.html" , SaveFormat.HTML);

			System.out.print("generateXml() setting xmlOutputStream to sectionData");

		}
	}

	public static void checkLicence() throws Exception {
		License license = new License();
		try {
			license.setLicense(Constants.LICENSE_DIR);
			System.out.print("Aspose License is Set!");

		} catch (Exception e) {
			System.out.print("Error in Aspose Licence check {}" + e);
			throw e;
		}
	}

	public static ArrayList extractContent(Node startNode, Node endNode, boolean isInclusive) throws Exception {
		// First check that the nodes passed to this method are valid for use.

		System.out.print("extractContent startNode {}" + startNode);
		System.out.print("extractContent endNode {}" + endNode);
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
				System.out.print("THERE IS SECTION HEADING WITHOUT NAME [{}{}]" + currNode.getText());
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

		System.out.print("verifyParameterNodes");
		System.out.print("verifyParameterNodes startNode {}" + startNode.getNodeType());
		System.out.print("verifyParameterNodes endNode {}" + endNode.getNodeType());

		System.out.print("verifyParameterNodes startNode.getAncestor {}" + startNode.getAncestor(NodeType.BODY));

		System.out.print("verifyParameterNodes endNode.getAncestor {}" + endNode.getAncestor(NodeType.BODY));
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
