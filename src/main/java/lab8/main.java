package lab8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;


public class main {

    public static void main(String[] args) {
        String inputFilePath = "src/main/java/lab8/laborator8_input.xlsx";
        String outputFilePath2 = "src/main/java/lab8/laborator8_output2.xlsx";
        String outputFilePath3 = "src/main/java/lab8/laborator8_output3.xlsx";

        try {
            System.out.println("=== 8.5.1 ");
            citireSiAfisareExcel(inputFilePath);

            System.out.println("\n=== 8.5.2");
            genereazaExcelCuMedieCalculata(inputFilePath, outputFilePath2);

            System.out.println("\n=== 8.5.3 ");
            genereazaExcelCuMedieFormula(inputFilePath, outputFilePath3);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void citireSiAfisareExcel(String inputFilePath) throws IOException {
        FileInputStream file = new FileInputStream(new File(inputFilePath));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getCellType()) {
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t\t");
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue() + "\t\t");
                        break;
                    default:
                        System.out.print("\t\t");
                }
            }
            System.out.println(" ");
        }
        workbook.close();
        file.close();
    }


    public static void genereazaExcelCuMedieCalculata(String inputFilePath, String outputFilePath) throws IOException {


        Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();

        FileInputStream file = new FileInputStream(new File(inputFilePath));
        XSSFWorkbook inputWorkbook = new XSSFWorkbook(file);
        XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);

        Iterator<Row> rowIterator = inputSheet.iterator();
        int idRand = 0;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            List<Object> rowData = new ArrayList<>();

            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case NUMERIC:
                        rowData.add(cell.getNumericCellValue());
                        break;
                    case STRING:
                        rowData.add(cell.getStringCellValue());
                        break;
                    default:
                        rowData.add("");
                }
            }

            if (idRand == 0) {

                rowData.add("Media");
            } else {
                if (rowData.size() >= 6) {
                    double nota1 = (Double) rowData.get(3);
                    double nota2 = (Double) rowData.get(4);
                    double nota3 = (Double) rowData.get(5);
                    double media = (nota1 + nota2 + nota3) / 3.0;

                    rowData.add(media);
                } else {
                    rowData.add(0.0);
                }
            }

            data.put(idRand, rowData.toArray());
            idRand++;
        }

        inputWorkbook.close();
        file.close();


        XSSFWorkbook workbook = new XSSFWorkbook(); // Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        Set<Integer> keyset = data.keySet();
        int rownum = 0;

        for (Integer key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;

            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);

                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);
                } else if (obj instanceof Double) {
                    cell.setCellValue((Double) obj);
                }
            }
        }

        try {
            FileOutputStream out = new FileOutputStream(new File(outputFilePath));
            workbook.write(out);
            out.close();
            System.out.println(outputFilePath + " written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }
        workbook.close();
    }
    public static void genereazaExcelCuMedieFormula(String inputFilePath, String outputFilePath) throws IOException {
        FileInputStream file = new FileInputStream(new File(inputFilePath));
        XSSFWorkbook inputWorkbook = new XSSFWorkbook(file);
        XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Calculate Formula");

        int rowIndex = 0;
        for (Row inputRow : inputSheet) {
            Row row = sheet.createRow(rowIndex);

            int cellnum = 0;
            for (Cell inputCell : inputRow) {
                Cell cell = row.createCell(cellnum++);

                switch (inputCell.getCellType()) {
                    case STRING:
                        cell.setCellValue(inputCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        cell.setCellValue(inputCell.getNumericCellValue());
                        break;
                    default:
                        cell.setCellValue("");
                }
            }

            if (rowIndex == 0) {
                row.createCell(cellnum).setCellValue("Media (Formula)");
            } else {
                int excelRow = rowIndex + 1;

                String formula = "AVERAGE(D" + excelRow + ":F" + excelRow + ")";
                row.createCell(cellnum).setCellFormula(formula);
            }
            rowIndex++;
        }

        inputWorkbook.close();
        file.close();

        try {
            FileOutputStream out = new FileOutputStream(new File(outputFilePath));
            workbook.write(out);
            out.close();
            System.out.println("Excel with formula cells written successfully (" + outputFilePath + ")");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        workbook.close();
    }

}
