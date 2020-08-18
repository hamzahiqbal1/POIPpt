import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;

import java.io.*;

/*
SlideSetup is the class where the Powerpoint is crated, and various methods
for creating different types of slides are contained.

 */

public class SlideSetup {
// The File object where the name of the output file is passed
    private File filename;
    // the PowerPoint representation
    private XMLSlideShow ppt;
    //Processor for reading input files
    private FileProcessor processor;

    /**
     *Constructor for Class SlideSetup
     *
     * @param name is the name of the file the powerpoint should be outputted to (e.g example.pptx)
     * @throws IOException
     * @throws XmlException
     */

    public SlideSetup(String name) throws IOException, XmlException {
    this.filename = new File(name);
    this.ppt = new XMLSlideShow();
    this.processor = new FileProcessor();
    }

    /**
     * Method for reading the data from an image
     * @param filename is the path for the image which needs to be read
     * @return image information in a form which can be used to add it to the Powerpoint
     * @throws IOException
     */

    public byte[] readImage(String filename) throws IOException {

        File image = new File(filename);
        byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
        return picture;
    }

    /**
     * Creates a Title slide using a default slide layout
     * @param title is what the main heading of the slide should read
     * @param sub is the subheading
     */

    public void createTitleSlide(String title, String sub) {


        //set up slide master
        XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);
        XSLFSlideLayout slideLayoutTitle = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

        //create slide using set layout
        XSLFSlide titleSlide = ppt.createSlide(slideLayoutTitle);
        //edit placeholder text
        XSLFTextShape title1 = titleSlide.getPlaceholder(0);
        title1.setText(title);

        //Edit subheading text
        XSLFTextShape subHeading = titleSlide.getPlaceholder(1);
        subHeading.setText(sub).setFontSize(12.0);


    }

    /**
     * Creates a slide contains an image and a body of text, along with a title. All which can be specified
     * @param filePath the path for an input file which is then displayed
     * @param imagePath the path for a .jpg file which is then displayed
     * @param title the main title for the slide
     * @throws IOException
     */


    public void createContentSlide(String filePath, String imagePath,String title) throws IOException {
        //set up slide master
        XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);
        XSLFSlideLayout slideLayoutTitle = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

        //create slide using set layout
        XSLFSlide slide = ppt.createSlide(slideLayoutTitle);
        //edit placeholder text
        XSLFTextShape title1 = slide.getPlaceholder(0);
        title1.setText(title);

        //Edit subheading text
        XSLFTextShape subHeading = slide.getPlaceholder(1);

        subHeading.setText(processor.readFile(filePath));

        XSLFPictureData idx = ppt.addPicture(readImage(imagePath),PictureData.PictureType.PNG );
        XSLFPictureShape pic = slide.createPicture(idx);


    }


    public void sendToFile() throws IOException {
        FileOutputStream out = new FileOutputStream(filename);

        ppt.write(out);
        System.out.println("You have successfully created a powerpoint");

    }


}
