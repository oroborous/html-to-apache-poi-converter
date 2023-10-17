package com.ivaalsolutions;

import org.apache.poi.sl.usermodel.AutoNumberingScheme;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;

import java.util.List;

public class HtmlPoiConverter {
    // Root-em sizes for heading tags
    private static final double H1_REM = 2.125;
    private static final double H2_REM = 1.875;
    private static final double H3_REM = 1.5;

    private static final List<String> PARAGRAPH_TAGS =
            List.of("p", "b", "i", "u", "li", "strong", "em", "h1", "h2", "h3");
    private static final List<String> RUN_TAGS = List.of("b", "i", "u", "strong", "em", "h1", "h2", "h3");

    public static void convertToPowerPoint(Document domTree, XSLFTextShape parentShape) {
        recursiveDFS(domTree.root(), parentShape);
    }

    private static void recursiveDFS(Node node, XSLFTextShape parentShape) {
        // We have encountered a text node outside of a paragraph element
        if (node instanceof TextNode textNode) {
            if (textNode.text().trim().length() > 0) {
                System.out.println("Text: " + textNode.text());
                // Create the paragraph and run and set their text
                XSLFTextParagraph p = parentShape.addNewTextParagraph();
                p.setBullet(false);
                XSLFTextRun r = p.addNewTextRun();
                r.setText(textNode.text());
            }
        } else if (node instanceof Element element) {
            // System.out.println("Element: " + element.toString());
            String tagName = element.tagName().toLowerCase();
            boolean makeParagraph = PARAGRAPH_TAGS.contains(tagName);

            // Make a new PowerPoint paragraph element
            if (makeParagraph) {
                // Text in paragraphs is bulleted by default. Only
                // keep the default bulleting for list items.
                boolean bulleted = tagName.equals("li");
                boolean numbered = bulleted && element.parent() != null &&
                        element.parent().tagName().equals("ol");

                XSLFTextParagraph p = parentShape.addNewTextParagraph();
                p.setBullet(bulleted);
                if (numbered) {
                    p.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 1);
                }

                // TODO: Don't we always need a run here?

                // Continue processing inner tags/text of the run
                for (Node child : element.childNodes()) {
                    // Does this tag require a text run because it's
                    // a formatting choice that cannot be applied at the
                    // paragraph level?
                    boolean makeRun = RUN_TAGS.contains(tagName);
                    if (makeRun) {
                        XSLFTextRun r = p.addNewTextRun();

                        // Only ever toggle a formatting to true, never false.
                        // e.g. Don't use r.setBold(tagName.equals("b"));
                        // Otherwise you might undo formatting performed by
                        // an outer, enclosing tag.
                        switch (tagName) {
                            case "b", "strong" -> r.setBold(true);
                            case "i", "em" -> r.setItalic(true);
                            case "u" -> r.setUnderlined(true);
                        }

                        recursiveDFS(child, r);
                    }
                    // Remove this element's children so they are not
                    // double-processed by any outer recursive loop
                    element.empty();
                }
            }
        }

        // Continue recursively processing remaining nodes
        for (Node child : node.childNodes()) {
            System.out.println(child.getClass() + ": " + child.nodeName());
            recursiveDFS(child, parentShape);
        }
    }

    private static void recursiveDFS(Node node, XSLFTextParagraph parentParagraph) {
        if (node instanceof TextNode textNode) {
            // We have encountered a text node while having
            // a paragraph element containing no text run
            if (textNode.text().trim().length() > 0) {
                System.out.println("Text: " + textNode.text());
                // Create the text run and set its text
                XSLFTextRun r = parentParagraph.addNewTextRun();
                r.setText(textNode.text());
            }
        } else if (node instanceof Element element) {
            String tagName = element.tagName().toLowerCase();
            boolean makeRun = RUN_TAGS.contains(tagName);

            if (makeRun) {
                for (Node child : element.childNodes()) {
                    XSLFTextRun r = parentParagraph.addNewTextRun();
                    switch (tagName) {
                        case "b", "strong" -> r.setBold(true);
                        case "i", "em" -> r.setItalic(true);
                        case "u" -> r.setUnderlined(true);
                        case "h1" -> r.setFontSize(r.getFontSize() * H1_REM);
                        case "h2" -> r.setFontSize(r.getFontSize() * H2_REM);
                        case "h3" -> r.setFontSize(r.getFontSize() * H3_REM);
                    }
                    recursiveDFS(child, r);
                }
                element.empty();
            }
        }
    }

    private static void recursiveDFS(Node node, XSLFTextRun parentRun) {
        // We have encountered a text node inside a run
        if (node instanceof TextNode textNode) {
            if (textNode.text().trim().length() > 0) {
                // Set the text
                parentRun.setText(textNode.text());
            }
        } else if (node instanceof Element element) {
            String tagName = element.tagName().toLowerCase();
            boolean continueRun = RUN_TAGS.contains(tagName);

            // More nested tags to continue formatting the run
            // we're working in
            if (continueRun) {
                switch (tagName) {
                    case "b", "strong" -> parentRun.setBold(true);
                    case "i", "em" -> parentRun.setItalic(true);
                    case "u" -> parentRun.setUnderlined(true);
                }
            }

            for (Node child : element.childNodes()) {
                recursiveDFS(child, parentRun);
            }

        }
    }
}
