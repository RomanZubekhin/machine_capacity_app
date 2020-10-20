package sample.controller;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class MachineCapacityController {

    @FXML
    private TextField planTextField;

    @FXML
    private TextField normTextField;

    @FXML
    private Button calculateButton;

    @FXML
    private Button openPathButtonPlan;

    @FXML
    private Button openPathButtonNorm;

    @FXML
    private TextArea areaTextField;

    String plan = "";
    String norm = "";

    @FXML
    void initialize() {
        openPathButtonPlan.setOnAction(event -> {
            openFilePlan(event);
        });
        openPathButtonNorm.setOnAction(event -> {
            openFileNorm(event);
        });

        calculateButton.setOnAction(event -> {
            try {
                final File fileRead = new File(plan);
                final File fileWrite = new File(norm);
                readFromExelAndWriteHashMap(fileRead,0,1);

                if (checkArticleInExel(fileWrite, hashMap)) {
                    areaTextField.appendText("Проверка выполнена!\n");
                    areaTextField.appendText("Запись данных...\n");
                    writeIntoExelNorm(fileWrite, hashMap);
                    areaTextField.appendText("Готово!");
                }else{
                    areaTextField.appendText("Запись данных не возможна! \nВнесите в таблицу следующие номера:");
                    for (String s : arrayList) {
                        areaTextField.appendText(s);
                    }
                }
            } catch (FileNotFoundException f){
                f.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
    }

    @FXML
    private void openFilePlan(ActionEvent event) {
        Stage stage = new Stage();
        FileChooser fileChooser = new FileChooser();//Класс работы с диалогом выборки и сохранения
        fileChooser.setTitle("Open Document");//Заголовок диалога
        FileChooser.ExtensionFilter extFilter =
                new FileChooser.ExtensionFilter("Exel files (*.xlsx)", "*.xlsx");//Расширение
        fileChooser.getExtensionFilters().add(extFilter);
        File file = fileChooser.showOpenDialog(stage);//Указываем текущую сцену
        if (file != null){
            planTextField.clear();
            planTextField.appendText(file.getPath());
            plan = planTextField.getText();
        }
    }
    @FXML
    private void openFileNorm(ActionEvent event) {
        Stage stage = new Stage();
        FileChooser fileChooser = new FileChooser();//Класс работы с диалогом выборки и сохранения
        fileChooser.setTitle("Open Document");//Заголовок диалога
        FileChooser.ExtensionFilter extFilter =
                new FileChooser.ExtensionFilter("Exel files (*.xlsx)", "*.xlsx");//Расширение
        fileChooser.getExtensionFilters().add(extFilter);
        File file = fileChooser.showOpenDialog(stage);//Указываем текущую сцену
        if (file != null){
            normTextField.clear();
            normTextField.appendText(file.getPath());
            norm = normTextField.getText();
        }
    }
    private final static HashMap<String,Integer> hashMap = new HashMap<>();
    private final static ArrayList<String> arrayList = new ArrayList<>();

    public static void readFromExelAndWriteHashMap(File file, int numArticle, int numQuantity) throws IOException {
        int count = 0;
        Workbook myExelBook = new XSSFWorkbook(new FileInputStream(file));

        Sheet myExelSheet = myExelBook.getSheet("Лист2");
        while (true) {
            try {
                Row row = myExelSheet.getRow(count);
                String article = null;
                int quantity = 0;
                Cell c = row.getCell(numArticle);
                if (!(c == null || c.getCellType() == CellType.BLANK)) {
                    if (row.getCell(numArticle).getCellType() == CellType.STRING) {
                        article = row.getCell(numArticle).getStringCellValue();
                    }
                    if (row.getCell(numQuantity).getCellType() == CellType.NUMERIC) {
                        quantity = (int) row.getCell(numQuantity).getNumericCellValue();
                    }
                    hashMap.put(article, quantity);
                    count++;
                } else break;
            }catch (NullPointerException exception){return;}
        }
    }

    public static boolean checkArticleInExel(File fileWrite, HashMap<String, Integer> map){
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
                    arrayList.add(m.getKey());
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

