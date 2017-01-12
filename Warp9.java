package effort;

//  UPDATE NOTES
/*  
   V1.0
    Gold copy of Warp9 is ready for use.
    Warp9 is the "main" file for XLR8, launch the main from this file.

   V1.1
    Fixed minor formatting issues.
    Improved error messages for missing input.
    Added comments to code.

   V1.2
    Moved all print outs to text area.
    Improved color scheme for buttons to NOT accomodate color blindness
    Minor format change in ExcelTouch.java

   V1.3
    "Choose File" button now shows file name instead of file path
    Fixed output message in ExcelTouch.java to show only the requested information
 */
import java.io.File; // all dem libs
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuBar;
import javafx.scene.control.MenuItem;
import javafx.scene.control.Separator;
import javafx.scene.control.TextArea;
import javafx.scene.effect.DropShadow;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

/**
 *
 * @author Kim Montes Reiner Liman
 */
public class Warp9 extends Application {

    String temp = "";
    ArrayList<String> appType = new ArrayList<>();
    ArrayList<String> campuses = new ArrayList<>();
    ArrayList<String> subjects = new ArrayList<>();
    ArrayList<CheckBox> appBtnArr = new ArrayList<>();
    ArrayList<CheckBox> campBtnArr = new ArrayList<>();
    ArrayList<CheckBox> subBtnArr = new ArrayList<>();

    @Override
    public void start(Stage primaryStage) {
        ExcelTouch et = new ExcelTouch(); // brand new everything
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        BorderPane root = new BorderPane();
        VBox checkBoxes = new VBox();
        GridPane mainContent = new GridPane();
        FileChooser fc = new FileChooser();
        DropShadow ds = new DropShadow();
        DatePicker sdp = new DatePicker();
        DatePicker edp = new DatePicker();
        TextArea info = new TextArea();
        MenuBar mb = new MenuBar();
        Menu help = new Menu("File");
        MenuItem about = new MenuItem("About");
        MenuItem reset = new MenuItem("Reset");
        Button fileBtn = new Button("Choose File");
        Button submit = new Button("Search");
//        Label type = new Label("Type");
//        Label campus = new Label("Campus");
//        Label subject = new Label("Subject");
//        CheckBox appt = new CheckBox("Appointment");
//        CheckBox drop = new CheckBox("Drop-ins");
//        CheckBox email = new CheckBox("E-mail");
//        CheckBox hmc = new CheckBox("HMC");
//        CheckBox davis = new CheckBox("Davis");
//        CheckBox traf = new CheckBox("Trafalgar");
//        CheckBox online = new CheckBox("Online");
//        CheckBox math = new CheckBox("Math");
//        CheckBox eng = new CheckBox("English");
//        CheckBox compSci = new CheckBox("Computer Programming");
//        CheckBox bizMath = new CheckBox("Business Math");
//        CheckBox onEng = new CheckBox("Online English");

        info.setFont(Font.font("Verdana", FontWeight.NORMAL, 16));

        help.getItems().addAll(reset, about); // menu bar, help == "File" on menu bar
        mb.getMenus().addAll(help);
        about.setOnAction(event -> {
            info.setText("Version 1.3\nLast Modified: 4 November 2016\n\nDate Created: 31 October 2016\nCreated by Reiner Liman and Kim Montes\n\nE.F.F.O.R.T\nExcel Formatting For Our Restless TAs\n\nSee Warp9.java for update notes");
        });

        sdp.setOnAction(event -> { // starting and ending date picker
            LocalDate startDate = sdp.getValue();
        });
        edp.setOnAction(event -> {
            LocalDate endDate = edp.getValue();
        });

        info.setEditable(false); // big text area in the middle
        fc.setTitle("Open Excel File"); // button for "Choose File" and the file chooser default
        fc.setInitialDirectory(new File(System.getProperty("user.home")));
        fc.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("All Files", "*.*"), new FileChooser.ExtensionFilter("Microsoft Excel Spreadsheet (.xls)", "*.xls"));

        fileBtn.addEventHandler(MouseEvent.MOUSE_ENTERED, event -> { // fancy effects for a couple buttons
            fileBtn.setEffect(ds);
        });
        submit.addEventHandler(MouseEvent.MOUSE_ENTERED, event -> {
            submit.setEffect(ds);
        });
        submit.addEventHandler(MouseEvent.MOUSE_EXITED, event -> {
            submit.setEffect(null);
        });
        fileBtn.addEventHandler(MouseEvent.MOUSE_EXITED, event -> {
            fileBtn.setEffect(null);
        });

        fileBtn.setOnAction(event -> { // change button text to file path once chosen and color it
            File file = fc.showOpenDialog(primaryStage);
            temp = file.toString();
            fileBtn.setText(file.getName());
            if (file.getName().substring(file.getName().lastIndexOf('.') + 1).equals("xls")) {
                fileBtn.setStyle("-fx-base: #71EEB8"); // alt color: 71EEB8 / 00FF00 (not color blind friendly)
                et.setInputFile(temp, this); // if right .xls file, send file to ExcelTouch
            } else {
                fileBtn.setStyle("-fx-base: #FF0000;");
            }
            for (int i = 0; i < appType.size(); i++) {
                appBtnArr.add(new CheckBox(this.appType.get(i)));
                appBtnArr.get(i).setSelected(false);
                checkBoxes.getChildren().add(this.appBtnArr.get(i));
            }
            checkBoxes.getChildren().add(new Separator());
            for (int i = 0; i < campuses.size(); i++) {
                campBtnArr.add(new CheckBox(this.campuses.get(i)));
                campBtnArr.get(i).setSelected(false);
                checkBoxes.getChildren().add(this.campBtnArr.get(i));
            }
            checkBoxes.getChildren().add(new Separator());
            for (int i = 0; i < subjects.size(); i++) {
                subBtnArr.add(new CheckBox(subjects.get(i)));
                subBtnArr.get(i).setSelected(false);
                checkBoxes.getChildren().add(this.subBtnArr.get(i));
            }
            checkBoxes.getChildren().add(new Separator());

        });

        reset.setOnAction(event -> { // clear er'rything on reset click
            System.out.println(checkBoxes.getChildren().size());
            //for (int i = 0; i < 16; i++) {
            checkBoxes.getChildren().clear();
            //}
//            appt.setSelected(false);
//            drop.setSelected(false);
//            email.setSelected(false);
//            traf.setSelected(false);
//            hmc.setSelected(false);
//            davis.setSelected(false);
//            online.setSelected(false);
//            eng.setSelected(false);
//            math.setSelected(false);
//            compSci.setSelected(false);
//            onEng.setSelected(false);
//            bizMath.setSelected(false);
            info.setText("");
            fileBtn.setText("Choose File");
            fileBtn.setTextFill(Color.rgb(0, 0, 0, 1));
            fileBtn.setStyle(null);
            sdp.setValue(null);
            edp.setValue(null);

//            for (int i = 0; i < this.appBtnArr.size(); i++) {
//                appBtnArr.get(i).setSelected(false);
//            }
//            for (int i = 0; i < this.campBtnArr.size(); i++) {
//                campBtnArr.get(i).setSelected(false);
//            }
//            for (int i = 0; i < this.subBtnArr.size(); i++) {
//                subBtnArr.get(i).setSelected(false);
//            }
        });

        submit.setOnAction(event -> { // submit your inputs vvv More details below
            boolean err = false;
            if (temp.substring(temp.lastIndexOf('.') + 1).equals("xls")) {
//                et.setInputFile(temp); // if right .xls file, send file to ExcelTouch
                Date start = new Date();
                Date end = new Date();
                info.setText("");
                try {
                    start = dateFormat.parse(sdp.getValue().toString()); // convert those dates to the proper format and as Date object
                    end = dateFormat.parse(edp.getValue().toString());
                } catch (Exception e) {
                    info.setText("Start date or end date missing.\n");
                    //e.getMessage();
                    //e.printStackTrace();
                }
                if (start.compareTo(end) > 0 || end.compareTo(start) < 0) { // dates don't make sense
                    info.setText("Invalid Dates.\n");
                    err = true;
                }
                int counter = 0;
                for (int i = 0; i < this.appBtnArr.size(); i++) {
                    if (appBtnArr.get(i).isSelected() == false) {
                        counter++;
                    }
                    if (counter == this.appBtnArr.size()) {
                        info.setText(info.getText() + "Please Select at Least One Type of Visit.\n");
                        err = true;
                    }
                }
                counter = 0;
                for (int i = 0; i < this.campBtnArr.size(); i++) {
                    if (campBtnArr.get(i).isSelected() == false) {
                        counter++;
                    }
                    if (counter == this.campBtnArr.size()) {
                        info.setText(info.getText() + "Please Select at Least One campus.\n");
                        err = true;
                    }
                }
                counter = 0;
                for (int i = 0; i < this.subBtnArr.size(); i++) {
                    if (subBtnArr.get(i).isSelected() == false) {
                        counter++;
                    }
                    if (counter == this.subBtnArr.size()) {
                        info.setText(info.getText() + "Please Select at Least One subject.\n");
                        err = true;
                    }
                }
                ArrayList<String> appTemp = new ArrayList<>();
                ArrayList<String> campTemp = new ArrayList<>();
                ArrayList<String> subTemp = new ArrayList<>();
                if (!err) {
                    for (int i = 0; i < this.appBtnArr.size(); i++) {
                        if (appBtnArr.get(i).isSelected() == true) {
                            appTemp.add(appBtnArr.get(i).getText());
                        }
                    }
                    for (int i = 0; i < this.campBtnArr.size(); i++) {
                        if (this.campBtnArr.get(i).isSelected() == true) {
                            campTemp.add(campBtnArr.get(i).getText());
                        }
                    }
                    for (int i = 0; i < this.subBtnArr.size(); i++) {
                        if (this.subBtnArr.get(i).isSelected() == true) {
                            subTemp.add(subBtnArr.get(i).getText());
                        }
                    }
                    try {
                        et.getInfo3(start, end, appTemp, campTemp, subTemp);
                    } catch (ParseException ex) {
                        Logger.getLogger(Warp9.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    info.setText(et.output);
//                    int appointment = (appt.isSelected()) ? 1 : 0;
//                    int dropin = (drop.isSelected()) ? 1 : 0;
//                    int emailBox = (email.isSelected()) ? 1 : 0;
//                    int campDavis = (davis.isSelected()) ? 1 : 0;
//                    int campTRA = (traf.isSelected()) ? 1 : 0;
//                    int campHMC = (hmc.isSelected()) ? 1 : 0;
//                    int campOnline = (online.isSelected()) ? 1 : 0;
//                    int subComp = (compSci.isSelected()) ? 1 : 0;
//                    int subEnglish = (eng.isSelected()) ? 1 : 0;
//                    int subMath = (math.isSelected()) ? 1 : 0;
//                    int subBusinessMath = (bizMath.isSelected()) ? 1 : 0;
//                    int subOnlineEnglish = (onEng.isSelected()) ? 1 : 0;
//                    String query = appointment + "" + dropin + "" + emailBox + "" + campDavis + "" + campTRA + "" + campHMC + "" + campOnline + "";
//                    query += subComp + "" + subEnglish + "" + subMath + "" + subBusinessMath + "" + subOnlineEnglish;
//                    et.getInfo2(start, end, query); // take options selected, send to ExcelTouch, and get the return output
//                    info.setText(et.output);
//                    System.out.println(appBtnArr.get(0).isSelected());
                }
            } else {
                info.setText("");
                info.setText("Please select an .xls file first.\n"); // bruh, why did you even open the program if you don't have an .xls file ready?
            }
        });

//        type.setFont(Font.font("Helvetica", FontWeight.BOLD, 16)); // fancy fonts for fancy headings
//        campus.setFont(Font.font("Helvetica", FontWeight.BOLD, 16));
//        subject.setFont(Font.font("Helvetica", FontWeight.BOLD, 16));
        checkBoxes.setAlignment(Pos.CENTER_LEFT); // rearranging checkboxes
        checkBoxes.setSpacing(5);
        checkBoxes.setPadding(new Insets(10, 10, 10, 10));
        //checkBoxes.getChildren().addAll(type, appt, drop, email, new Separator(), campus, hmc, davis, traf, online, new Separator(), subject, math, eng, compSci, bizMath, onEng);
//        checkBoxes.getChildren().add(this.appBtnArr.get(0));
        info.setPrefWidth(820); // info box size
        info.setPrefHeight(360);

        mainContent.setHgap(10); // where to put all of the elements to the right of checkboxes by grid
        mainContent.setVgap(10);
        mainContent.setPadding(new Insets(10, 10, 10, 10));
        mainContent.add(fileBtn, 0, 0, 4, 1);
        mainContent.add(sdp, 4, 0);
        mainContent.add(edp, 5, 0);
        mainContent.add(submit, 6, 0);
        mainContent.add(info, 0, 1, 7, 7);

        root.setLeft(checkBoxes); // add all to root, root is borderPane
        root.setCenter(mainContent);
        root.setTop(mb);
        Scene scene = new Scene(root, 1080, 480); // size of application window

        primaryStage.setTitle("EFFORT"); // title, set scene, and (ACTION!) show the program running
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public void setArrays(ArrayList<String> appType, ArrayList<String> campusArr, ArrayList<String> subjectArr) {
        this.appType = appType;
        this.campuses = campusArr;
        this.subjects = subjectArr;
    }
} // Warp9
