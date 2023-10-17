package com.ivaalsolutions;

import org.apache.poi.xslf.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class Main {
    public static void main(String[] args) {

        try (XMLSlideShow ppt = new XMLSlideShow();
             InputStream input = ClassLoader.getSystemResourceAsStream("input.html")
        ) {
            Document doc = Jsoup.parse(input, "UTF-8", "");

            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
            XSLFSlideLayout titleBodyLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
            XSLFSlide slide = ppt.createSlide(titleBodyLayout);
            XSLFTextShape slideBody = slide.getPlaceholder(1);
            slideBody.clearText();

            HtmlPoiConverter.convertToPowerPoint(doc, slideBody);

            ppt.write(new FileOutputStream("output.pptx"));
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }
}