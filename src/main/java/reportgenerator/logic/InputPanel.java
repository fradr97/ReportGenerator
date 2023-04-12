package reportgenerator.logic;

import javax.swing.*;
import java.io.IOException;

public class InputPanel {

    public void showInputPanel() throws IOException {
        JTextField inputFolderPath = new JTextField();
        JTextField outputFilePath = new JTextField();
        JTextField fileType = new JTextField();
        Object[] message = {
                "Input folder path:", inputFolderPath,
                "Output file path:", outputFilePath,
                "File type (e - enum, r - repr/req, m - model):", fileType
        };

        int option = JOptionPane.showConfirmDialog(null, message, "Report Generator", JOptionPane.OK_CANCEL_OPTION);
        if (option == JOptionPane.OK_OPTION) {
            if (!inputFolderPath.getText().equals("") &&
                    !outputFilePath.getText().equals("") &&
                    !fileType.getText().equals("") &&
                    (fileType.getText().equals("e") || fileType.getText().equals("r") || fileType.getText().equals("m"))) {

                GenerateDoc generateDoc = new GenerateDoc();
                generateDoc.genDoc(inputFolderPath.getText().trim(), outputFilePath.getText().trim(), fileType.getText().trim());
            }
        }
    }
}
