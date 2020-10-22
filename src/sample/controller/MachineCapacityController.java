package sample.controller;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import sample.model.ClassMethods;

import java.io.*;

public class MachineCapacityController extends ClassMethods {

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
        openPathButtonPlan.setOnAction(this::openFilePlan);
        openPathButtonNorm.setOnAction(this::openFileNorm);

        calculateButton.setOnAction(event -> {
            final File fileRead = new File(plan);
            final File fileWrite = new File(norm);
            readFromExelAndWriteHashMap(fileRead,hashMap);

            if (checkArticleInExel(fileWrite, hashMap,arrayList)) {
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
}

