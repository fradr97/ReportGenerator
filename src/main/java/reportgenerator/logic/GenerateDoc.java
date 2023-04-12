package reportgenerator.logic;

import reportgenerator.file.ReadFile;
import reportgenerator.file.WriteExcel;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.IOException;
import java.util.Objects;
import java.util.Scanner;

public class GenerateDoc {
    private static final String FIRST_SHEET_NAME = "Total";

    public GenerateDoc() {
    }

    public void genDoc(String inputFolderPath, String outputFilePath, String fileType) throws IOException {
        ReadFile reader = new ReadFile();
        File folder = new File(inputFolderPath);

        WriteExcel writeExcel = new WriteExcel();

        String example;
        String description;
        String type;
        String fieldName;

        switch (fileType) {
            case "e": // ENUM
                writeExcel.createEnumSheet(FIRST_SHEET_NAME);

                for (final File fileEntry : Objects.requireNonNull(folder.listFiles())) {
                    Scanner myReader = reader.readFile(folder + "\\" + fileEntry.getName());
                    writeExcel.createEnumSheet(fileEntry.getName());

                    int i = 1;
                    while (myReader.hasNextLine()) {
                        String data = myReader.nextLine();

                        if (data.trim().contains("*")) {
                            fieldName = StringUtils.substringBetween(data.trim(), "*", "-").trim();
                            description = StringUtils.substringBetween(data.trim(), "-", "\"")
                                    .replace("\\n", "").trim();

                            writeExcel.insertEnumRow(i, fileEntry.getName(), fieldName, description);
                            i++;
                        }
                    }
                }
                break;
            case "r": // INTERFACE - REPR/REQ
                writeExcel.createInterfaceSheet(FIRST_SHEET_NAME);

                for (final File fileEntry : Objects.requireNonNull(folder.listFiles())) {
                    Scanner myReader = reader.readFile(folder + "\\" + fileEntry.getName());
                    writeExcel.createInterfaceSheet(fileEntry.getName());

                    int i = 1;
                    while (myReader.hasNextLine()) {
                        example = "";
                        description = "";
                        type = "";
                        fieldName = "";

                        String data = myReader.nextLine();

                        if (data.trim().startsWith("@Schema") && !data.trim().startsWith("@Schema(name")) {
                            if (data.contains("example")) {
                                example = StringUtils.substringBetween(data, "example = \"", "\"");
                            }
                            if (data.contains("description")) {
                                description = StringUtils.substringBetween(data, "description = \"", "\"");
                            }

                            // type and field
                            try {
                                data = myReader.nextLine();

                                type = data.trim().split(" ")[0];
                                fieldName = data.trim().split(" ")[1].replace("();", "");
                            } catch (Exception ignored) {
                            }

                            writeExcel.insertInterfaceRow(i, fileEntry.getName(), fieldName, type, example, description);
                            i++;
                        }
                    }
                }
                break;
            case "m": // MODEL
                writeExcel.createModelSheet(FIRST_SHEET_NAME);

                for (final File fileEntry : Objects.requireNonNull(folder.listFiles())) {
                    Scanner myReader = reader.readFile(folder + "\\" + fileEntry.getName());
                    writeExcel.createModelSheet(fileEntry.getName());

                    int i = 1;
                    while (myReader.hasNextLine()) {
                        String data = myReader.nextLine();

                        if ((data.trim().startsWith("private")
                                || data.trim().startsWith("public")
                                || data.trim().startsWith("protected"))
                                && !data.trim().contains("class")) {
                            type = data.trim().split(" ")[1];
                            fieldName = data.trim().split(" ")[2].replace(";", "");

                            writeExcel.insertModelRow(i, fileEntry.getName(), type, fieldName);
                            i++;
                        }
                    }
                }
                break;
            default:
                break;
        }
        writeExcel.write(outputFilePath);
    }
}
