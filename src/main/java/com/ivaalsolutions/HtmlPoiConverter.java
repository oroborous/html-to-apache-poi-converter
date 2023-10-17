package com.ivaalsolutions;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;

public class HtmlPoiConverter {

    public static void convertToPowerPoint(Document domTree, XSLFTextShape parentShape) {
        recursiveDFS(domTree.root(), parentShape);
    }

    private static void recursiveDFS(Node node, XSLFTextShape parentShape) {
        if (node instanceof TextNode textNode) {
            if (textNode.text().trim().length() > 0) {
                System.out.println("Text: " + textNode.text());
                XSLFTextParagraph p = parentShape.addNewTextParagraph();
                p.setBullet(false);
                XSLFTextRun r = p.addNewTextRun();
                r.setText(textNode.text());
            }
        } else if (node instanceof Element element) {
            //            System.out.println("Element: " + element.toString());
            String tagName = element.tagName();
            if (tagName.equalsIgnoreCase("p")) {
                XSLFTextParagraph p = parentShape.addNewTextParagraph();
                p.setBullet(false);
                for (Node child : element.childNodes()) {
                    recursiveDFS(child, p);
                }
                element.empty();
            } else if (tagName.equalsIgnoreCase("b")) {
                XSLFTextParagraph p = parentShape.addNewTextParagraph();
                p.setBullet(false);
                XSLFTextRun r = p.addNewTextRun();
                r.setBold(true);
                for (Node child : element.childNodes()) {
                    recursiveDFS(child, r);
                }
                element.empty();
            } else if (tagName.equalsIgnoreCase("i")) {
                XSLFTextParagraph p = parentShape.addNewTextParagraph();
                p.setBullet(false);
                XSLFTextRun r = p.addNewTextRun();
                r.setItalic(true);
                for (Node child : element.childNodes()) {
                    recursiveDFS(child, r);
                }
                element.empty();
            }
        }
        for (Node child : node.childNodes()) {
            System.out.println(child.getClass() + ": " + child.nodeName());
            recursiveDFS(child, parentShape);
        }
    }

    private static void recursiveDFS(Node node, XSLFTextParagraph parentParagraph) {
        if (node instanceof TextNode textNode) {
            if (textNode.text().trim().length() > 0) {
                System.out.println("Text: " + textNode.text());
                XSLFTextRun r = parentParagraph.addNewTextRun();
                r.setText(textNode.text());
            }
        } else if (node instanceof Element element) {
            String tagName = element.tagName();
            if (tagName.equalsIgnoreCase("b")) {
                XSLFTextRun r = parentParagraph.addNewTextRun();
                r.setBold(true);
                for (Node child : element.childNodes()) {
                    recursiveDFS(child, r);
                }
                element.empty();
            } else if (tagName.equalsIgnoreCase("i")) {
                XSLFTextRun r = parentParagraph.addNewTextRun();
                r.setItalic(true);
                for (Node child : element.childNodes()) {
                    recursiveDFS(child, r);
                }
                element.empty();
            }
        }
    }

    private static void recursiveDFS(Node node, XSLFTextRun parentRun) {
        if (node instanceof TextNode textNode) {
            if (textNode.text().trim().length() > 0) {
                parentRun.setText(textNode.text());
            }
        } else if (node instanceof Element element) {
            String tagName = element.tagName();
            if (tagName.equalsIgnoreCase("b")) {
                parentRun.setBold(true);
                for (Node child : element.childNodes()) {
                    recursiveDFS(child, parentRun);
                }
            } else if (tagName.equalsIgnoreCase("i")) {
                parentRun.setItalic(true);
                for (Node child : element.childNodes()) {
                    recursiveDFS(child, parentRun);
                }
            }
        }
    }
}
