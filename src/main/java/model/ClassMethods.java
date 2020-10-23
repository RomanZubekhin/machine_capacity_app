package model;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class ClassMethods {
    protected static final HashMap<String,Integer> hashMap = new HashMap<>();
    protected static final ArrayList<String> arrayList = new ArrayList<>();

    public static void readFromExelAndWriteHashMap(File file, HashMap<String,Integer> map){
        Workbook myExelBook = null;
        try {
            myExelBook = new XSSFWorkbook(new FileInputStream(file));
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet myExelSheet = myExelBook.getSheetAt(1);
        String article = null;
        int quantity = 0;
        for (Row row : myExelSheet) {
            for (Cell c : row) {
                if (!(c == null || c.getCellType() == CellType.BLANK)) {
                    if (c.getCellType() == CellType.STRING) {
                        article = c.getStringCellValue();
                    }
                    if (c.getCellType() == CellType.NUMERIC) {
                        quantity = (int) c.getNumericCellValue();
                    }
                } else break;
            }
            map.put(article, quantity);
        }
    }

    public static boolean checkArticleInExel(File fileWrite, HashMap<String, Integer> map, ArrayList<String> array){
        boolean flagWrite = false;
        try {
            FileInputStream fileInputStream = new FileInputStream(fileWrite);
            Workbook wb = new XSSFWorkbook(fileInputStream);
            Sheet sheet = wb.getSheetAt(0);
            int startCell = 0;
            int sizeMap = map.size();
            int hitCounter = 0;
            for (Map.Entry<String, Integer> m : map.entrySet()) {
                boolean flag = true;
                for (Row row : sheet) {
                    DataFormatter df = new DataFormatter();
                    Cell cell = row.getCell(startCell);
                    String val = df.formatCellValue(cell);
                    if (m.getKey().equals(val)) {
                        hitCounter++;
                        flag = false;
                    }else if (val == null  || cell.getCellType() == CellType.BLANK){
                        break;
                    }
                }
                if(flag){
                    array.add(m.getKey());
                }
            }
            if (sizeMap == hitCounter){
                flagWrite = true;
            }
            fileInputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return flagWrite;
    }

    public static void writeIntoExelNorm(File fileWrite, HashMap<String, Integer> map){
        try {
            FileInputStream fileInputStream = new FileInputStream(fileWrite);
            Workbook wb = new XSSFWorkbook(fileInputStream);
            Sheet sheet = wb.getSheetAt(0);
            boolean flag = false;
            int startCell = 0;
            for (Map.Entry<String, Integer> m : map.entrySet()) {
                String article;
                int quantity;
                for (Row row : sheet) {
                    Cell cell = row.getCell(startCell);
                    article = m.getKey();
                    quantity = m.getValue();
                    if (article.equals(cell.getStringCellValue())) {
                        for (Cell c : row) {
                            if (flag) {
                                c.setCellValue(quantity);
                                flag = false;
                            } else if (c.getCellType() == CellType.BLANK || c.getCellType() == CellType.FORMULA) {
                                flag = false;
                            } else if (c.getCellType() == CellType.NUMERIC) {
                                flag = true;
                            } else if (c.getStringCellValue().equals("end")) {
                                break;
                            }
                        }
                    } else if (cell.getStringCellValue().equals("end")) {
                        break;
                    }
                }
            }

            //Re-evaluate formulas with POI's FormulaEvaluator
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();
            //write data
            FileOutputStream fileOutputStream = new FileOutputStream(fileWrite);
            wb.write(fileOutputStream);
            fileInputStream.close();
            fileOutputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
