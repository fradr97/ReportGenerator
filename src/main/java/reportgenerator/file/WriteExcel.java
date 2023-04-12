package reportgenerator.file;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {
    private static final int COLUMN_WIDTH = 8000;

    private Workbook workbook;
    private Sheet sheet;
    private CellStyle style;
    private Row row;

    public WriteExcel() {
        workbook = new XSSFWorkbook();
        style = workbook.createCellStyle();
        style.setWrapText(true);
    }

    /************************ ENUM ************************/
    public void createEnumSheet(String sheetName) {
        sheet = workbook.createSheet(sheetName);
        sheet.setColumnWidth(0, COLUMN_WIDTH);
        sheet.setColumnWidth(1, COLUMN_WIDTH);
        sheet.setColumnWidth(2, COLUMN_WIDTH);
        sheet.setColumnWidth(3, COLUMN_WIDTH);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Classe");
        headerCell.setCellStyle(headerStyle);

        CellStyle headerFieldStyle = workbook.createCellStyle();
        headerFieldStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        headerFieldStyle.setFillPattern(headerStyle.getFillPattern());
        headerFieldStyle.setAlignment(headerStyle.getAlignment());
        headerFieldStyle.setFont(font);
        headerCell = header.createCell(1);
        headerCell.setCellValue("Campo enum");
        headerCell.setCellStyle(headerFieldStyle);

        headerCell = header.createCell(2);
        headerCell.setCellValue("Descrizione");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(3);
        headerCell.setCellValue("Nome campo in AFU");
        headerCell.setCellStyle(headerStyle);
    }

    public void insertEnumRow(int rowIndex, String className, String field, String description) {
        row = sheet.createRow(rowIndex);

        Cell cell = row.createCell(0);
        cell.setCellValue(className);
        cell.setCellStyle(style);

        CellStyle fieldStyle = workbook.createCellStyle();
        fieldStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        fieldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell(1);
        cell.setCellValue(field);
        cell.setCellStyle(fieldStyle);

        cell = row.createCell(2);
        cell.setCellValue(description);
        cell.setCellStyle(style);

        cell = row.createCell(3);
        cell.setCellValue("");
        cell.setCellStyle(style);
    }

    /************************ INTERFACE ************************/
    public void createInterfaceSheet(String sheetName) {
        sheet = workbook.createSheet(sheetName);
        sheet.setColumnWidth(0, COLUMN_WIDTH);
        sheet.setColumnWidth(1, COLUMN_WIDTH);
        sheet.setColumnWidth(2, COLUMN_WIDTH);
        sheet.setColumnWidth(3, COLUMN_WIDTH);
        sheet.setColumnWidth(4, COLUMN_WIDTH);
        sheet.setColumnWidth(5, COLUMN_WIDTH);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Classe");
        headerCell.setCellStyle(headerStyle);

        CellStyle headerFieldStyle = workbook.createCellStyle();
        headerFieldStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        headerFieldStyle.setFillPattern(headerStyle.getFillPattern());
        headerFieldStyle.setAlignment(headerStyle.getAlignment());
        headerFieldStyle.setFont(font);
        headerCell = header.createCell(1);
        headerCell.setCellValue("Campo");
        headerCell.setCellStyle(headerFieldStyle);

        headerCell = header.createCell(2);
        headerCell.setCellValue("Tipo");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(3);
        headerCell.setCellValue("Esempio");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(4);
        headerCell.setCellValue("Descrizione");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(5);
        headerCell.setCellValue("Nome campo in AFU");
        headerCell.setCellStyle(headerStyle);
    }

    public void insertInterfaceRow(int rowIndex, String className, String field, String type, String example, String description) {
        row = sheet.createRow(rowIndex);

        Cell cell = row.createCell(0);
        cell.setCellValue(className);
        cell.setCellStyle(style);

        CellStyle fieldStyle = workbook.createCellStyle();
        fieldStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        fieldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell(1);
        cell.setCellValue(field);
        cell.setCellStyle(fieldStyle);

        cell = row.createCell(2);
        cell.setCellValue(type);
        cell.setCellStyle(style);

        cell = row.createCell(3);
        cell.setCellValue(example);
        cell.setCellStyle(style);

        cell = row.createCell(4);
        cell.setCellValue(description);
        cell.setCellStyle(style);

        cell = row.createCell(5);
        cell.setCellValue("");
        cell.setCellStyle(style);
    }

    /************************ MODEL ************************/
    public void createModelSheet(String sheetName) {
        sheet = workbook.createSheet(sheetName);
        sheet.setColumnWidth(0, COLUMN_WIDTH);
        sheet.setColumnWidth(1, COLUMN_WIDTH);
        sheet.setColumnWidth(2, COLUMN_WIDTH);
        sheet.setColumnWidth(3, COLUMN_WIDTH);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Classe");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Tipo");
        headerCell.setCellStyle(headerStyle);

        CellStyle headerFieldStyle = workbook.createCellStyle();
        headerFieldStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        headerFieldStyle.setFillPattern(headerStyle.getFillPattern());
        headerFieldStyle.setAlignment(headerStyle.getAlignment());
        headerFieldStyle.setFont(font);
        headerCell = header.createCell(2);
        headerCell.setCellValue("Campo");
        headerCell.setCellStyle(headerFieldStyle);

        CellStyle afuStyle = workbook.createCellStyle();
        afuStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        headerCell = header.createCell(3);
        headerCell.setCellValue("Nome campo in AFU");
        headerCell.setCellStyle(headerStyle);
    }

    public void insertModelRow(int rowIndex, String className, String type, String field) {
        row = sheet.createRow(rowIndex);

        Cell cell = row.createCell(0);
        cell.setCellValue(className);
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue(type);
        cell.setCellStyle(style);

        CellStyle fieldStyle = workbook.createCellStyle();
        fieldStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        fieldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell(2);
        cell.setCellValue(field);
        cell.setCellStyle(fieldStyle);

        cell = row.createCell(3);
        cell.setCellValue("");
        cell.setCellStyle(style);
    }

    public void write(String filePath) throws IOException {
        File path = new File(filePath);
        FileOutputStream outputStream = new FileOutputStream(path + ".xlsx");
        workbook.write(outputStream);
        workbook.close();
    }

//    public void readAllSheets() {
//        for(int j = 1; j < workbook.getNumberOfSheets(); j ++) {
//            Sheet sheet = workbook.getSheetAt(j);
//
//            int i = 0;
//            for (Row row : sheet) {
//                Row newRow = sheet.createRow(i);
//                int cellIndex = 0;
//                for (Cell cell : row) {
//                    switch (cell.getCellType()) {
//                        case STRING:
//                            Cell stringCell = newRow.createCell(cellIndex);
//                            System.out.println(cell.getRichStringCellValue().getString());
//                            stringCell.setCellValue(cell.getRichStringCellValue().getString());
//                            stringCell.setCellStyle(style);
//
////                            data.get(i).add(cell.getRichStringCellValue().getString());
//                            break;
//                        case NUMERIC:
//                            Cell numericCell = newRow.createCell(cellIndex);
//                            numericCell.setCellValue(String.valueOf(cell.getNumericCellValue()));
//                            numericCell.setCellStyle(style);
//
////                            data.get(i).add(String.valueOf(cell.getNumericCellValue()));
//                            break;
//                        case BOOLEAN:
//                            Cell boolCell = newRow.createCell(cellIndex);
//                            boolCell.setCellValue(String.valueOf(cell.getBooleanCellValue()));
//                            boolCell.setCellStyle(style);
//
////                            data.get(i).add(String.valueOf(cell.getBooleanCellValue()));
//                            break;
//                        case FORMULA:
//                            Cell formulaCell = newRow.createCell(cellIndex);
//                            formulaCell.setCellValue(cell.getCellFormula());
//                            formulaCell.setCellStyle(style);
//
////                            data.get(i).add(cell.getCellFormula());
//                            break;
//                        default:
//                            Cell defCell = newRow.createCell(cellIndex);
//                            defCell.setCellValue(" ");
//                            defCell.setCellStyle(style);
//
////                            data.get(i).add(" ");
//                    }
//                    cellIndex ++;
//                }
//                i++;
//            }
//        }
//    }

//    public void writeAll(Map<Integer, List<String>> map, Sheet sheet) {
//        sheet = workbook.getSheetAt(j);
//
//        int i = 0;
//        for (Row row : sheet) {
//
//
//        }
//        for(int i = 0; i < map.values().size(); i ++) {
//            row = sheet.createRow(i);
//
//            for(int j = 0; j < map.size(); j ++) {
//                Cell cell = row.createCell(j);
//                cell.setCellValue(map.get(i).get(j));
//                cell.setCellStyle(style);
//            }
//        }
//    }
}
