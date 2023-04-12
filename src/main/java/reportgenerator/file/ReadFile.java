package reportgenerator.file;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.Scanner;

public class ReadFile {

    public Scanner readFile(String filePath) {
        try {
            File myObj = new File(filePath);
            return new Scanner(myObj);
        } catch (FileNotFoundException e) {
            System.out.println("An error occurred.");
            e.printStackTrace();
        }
        return null;
    }
}
