package net.qvarford;

import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

public class Main {
    private static final Logger logger
            = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) throws IOException {
        logger.info("Hello INFO!");


        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().getFirst();
            XSLFSlideLayout layout
                    = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

            XSLFSlide slide = ppt.createSlide(layout);
            XSLFTextShape titleShape = slide.getPlaceholder(0);
            titleShape.clearText();
            titleShape.appendText("My title!", false);
            XSLFTextShape contentShape = slide.getPlaceholder(1);
            contentShape.clearText();
            contentShape.appendText("My content first", true);
            contentShape.appendText("My content again", true);
            contentShape.appendText("My content third", true);

            XSLFNotes note = ppt.getNotesSlide(slide);

            Arrays.stream(note.getPlaceholders())
                    .filter(shape -> shape.getTextType() == Placeholder.BODY)
                    .forEach(shape -> shape.setText("""
                            Hello
                            I'm a note.
                            Nice to meet you.
                            """));

            FileOutputStream out = new FileOutputStream("powerpoint.pptx");
            ppt.write(out);
            out.close();
        }
        System.out.println("Hello world!");
    }
}