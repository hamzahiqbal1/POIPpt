
import org.apache.xmlbeans.XmlException;

import java.io.IOException;
import java.util.Scanner;


public class Powerpoint {
    public static void main(String[] args) throws IOException, XmlException {
        Scanner obj = new Scanner(System.in);
        System.out.println("Enter the name of desired output file");

        String filename = obj.nextLine();

        SlideSetup presentation = new SlideSetup(filename);

        presentation.createTitleSlide("Memes","This powerpoint is about memes");
        presentation.createTitleSlide("ice", "is cold");
        presentation.createContentSlide( "C:\\Users\\ihamz\\POIresources\\test.txt","Original_Doge_meme.jpg","doge");
        presentation.sendToFile();


    }
}

