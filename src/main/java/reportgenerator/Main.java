package reportgenerator;

import reportgenerator.logic.InputPanel;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {
        InputPanel inputPanel = new InputPanel();
        inputPanel.showInputPanel();
    }
}
