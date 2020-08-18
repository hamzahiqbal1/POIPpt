import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class FileProcessor {
    public FileProcessor() {

    }

    public static String readFile(String path) throws IOException {
        String content = Files.lines(Paths.get(path), StandardCharsets.UTF_8)
                .collect(Collectors.joining(System.lineSeparator()));

        return content;
    }



}