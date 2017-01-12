package effort;

import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import javafx.scene.control.CheckBox;
import javax.swing.JOptionPane;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * @author Kim Montes Reiner Liman
 *
 * This will be interacting with the excel file. Only XLS files are accepted
 */
public class ExcelTouch {

    private String inputFile;//this is the excel file that this will be interacting with
    File inputWorkbook;
    Workbook w;
    String[][] infoArr; //array of entries
    SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yy");
    LinkedList<Entry> searchRes = new LinkedList<>();
    String typeSubString;
    String campusSubString;
    String subjectSubString;
    boolean removedOne = false;
    String queryIn = "";
    String output = "";
    ArrayList<String> startList = new ArrayList<>();
    int studentCount = 0;
    ArrayList<String> typeArr = new ArrayList<>();
    ArrayList<String> typeCamp = new ArrayList<>();
    ArrayList<String> typeSubject = new ArrayList<>();

    public void setInputFile(String inputFile, Warp9 temp) { //called to pass the excel file name over to ExcelTouch
        this.inputFile = inputFile;
        inputWorkbook = new File(inputFile);
        this.setArray(temp);
    }

    public void setArray(Warp9 temp) { //Getting all entries from the excel file
        try {
            w = Workbook.getWorkbook(inputWorkbook);
            Sheet sheet = w.getSheet(0); // Get the first sheet
            infoArr = new String[sheet.getRows()][6];
            Cell cell;
            for (int i = 1; i < sheet.getRows(); i++) { //gets the date of visit
                cell = sheet.getCell(0, i);
                infoArr[i - 1][0] = cell.getContents();
            }
            for (int i = 1; i < sheet.getRows(); i++) { //gets the Type of visit (Appointment/Drop in)
                cell = sheet.getCell(4, i);
                infoArr[i - 1][1] = cell.getContents();
            }
            for (int i = 1; i < sheet.getRows(); i++) { //gets the ammount of time of viist
                cell = sheet.getCell(8, i);
                infoArr[i - 1][2] = cell.getContents();
            }
            for (int i = 1; i < sheet.getRows(); i++) { //gets subject of visit
                cell = sheet.getCell(10, i);
                infoArr[i - 1][3] = cell.getContents();
            }
            for (int i = 1; i < sheet.getRows(); i++) { //gets campus 
                cell = sheet.getCell(12, i);
                infoArr[i - 1][4] = cell.getContents();
            }
            for (int i = 1; i < sheet.getRows(); i++) { //gets name of student(used for testing purposes)
                cell = sheet.getCell(1, i);
                infoArr[i - 1][5] = cell.getContents();
            }

            for (int i = 0; i < infoArr.length - 1; i++) {//Gets rid of all entries not inbetween the start and end date
                typeArr.add(infoArr[i][1]);
                typeCamp.add(infoArr[i][4]);
                typeSubject.add(infoArr[i][3]);
            }
            typeArr = removeDuplicates(typeArr);
            typeCamp = removeDuplicates(typeCamp);
            typeSubject = removeDuplicates(typeSubject);
            temp.setArrays(typeArr, typeCamp, typeSubject);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//setArray

    public void getInfo3(Date start2, Date end2, ArrayList<String> apps, ArrayList<String> camp, ArrayList<String> sub) throws ParseException {
        Date start = start2;
        Date end = end2;
        Date question; //this is assigned the value of the date from each entry from the excel file
        //int count = 0;
        //int placeHolder = 0;
        searchRes.clear();
        for (int i = 0; i < infoArr.length - 1; i++) {//Gets rid of all entries not inbetween the start and end date
            question = dateFormat.parse(infoArr[i][0]);
            if (question.compareTo(start) >= 0 && question.compareTo(end) <= 0) {
                //count++;
                Entry tempEntry = new Entry();//creates an temporary entry object
                tempEntry.setDate(infoArr[i][0]);
                tempEntry.setType(infoArr[i][1]);
                tempEntry.setTime(infoArr[i][2]);
                tempEntry.setSubject(infoArr[i][3]);
                tempEntry.setCampus(infoArr[i][4]);
                tempEntry.setName(infoArr[i][5]);

                typeArr.add(infoArr[i][1]);
                typeCamp.add(infoArr[i][4]);
                typeSubject.add(infoArr[i][3]);
                searchRes.add(tempEntry);//entry object added to searchRes for later use
            }
        }
        int j = 0;
        //for (int j = 0; j < searchRes.size(); j++) {
        while (j < searchRes.size()) {
            String search = "";
            String btnTxt = "";
            int counter = 0;
            int counter2 = 0;
            int counter3 = 0;
            for (int i = 0; i < apps.size(); i++) {
                search = searchRes.get(j).getType();
                btnTxt = apps.get(i);
                if (search.equals(btnTxt)) {
                    counter++;
                }
            }
            for (int i = 0; i < camp.size(); i++) {
                search = searchRes.get(j).getCampus();
                btnTxt = camp.get(i);
                if (search.equals(btnTxt)) {
                    counter2++;
                }
            }
            for (int i = 0; i < sub.size(); i++) {
                search = searchRes.get(j).getSubject();
                btnTxt = sub.get(i);
                if (search.equals(btnTxt)) {
                    counter3++;
                }
            }
            if (counter == 0 || counter2 == 0 || counter3 == 0) {
                searchRes.remove(j);
            } else {
                j++;
            }
        }
        //this.output = searchRes.size() + "!";
        for (int i = 0; i < searchRes.size(); i++) {
            System.out.println(searchRes.get(i).getDate() + " " + searchRes.get(i).getName() + " " + searchRes.get(i).getType() + " " + searchRes.get(i).getSubject() + " " + searchRes.get(i).getCampus());
        }
        startList.clear();
        for (int x = 0; x < searchRes.size(); x++) {
            startList.add(searchRes.get(x).getName());
        }
        Set<String> endList = new HashSet<String>(startList);
        studentCount = endList.size();
        this.outputWindow2(apps, camp, sub);
    }

    public void getInfo2(Date start2, Date end2, String query) { //called to get specific information 
        /**
         * start2: Earliest date that the program will use to look through end2:
         * Latest date that the program will use to look through query: Contains
         * a string of 1s and 0s that dictate what the user is looking for
         * Example: String: 0000000000 The first 2 0s represent Appointments or
         * Drop-ins The next 3 0s represent a campus The last 5 0s represent a
         * subject 0s mean they are not looking for information on that field 1s
         * mean they are looking for that field.
         */
        try {
            Date start = start2;
            Date end = end2;
            Date question; //this is assigned the value of the date from each entry from the excel file
            queryIn = query;
            typeSubString = query.substring(0, 3);
            campusSubString = query.substring(3, 6);
            subjectSubString = query.substring(6, 11);
            //int count = 0;
            //int placeHolder = 0;
            searchRes.clear();
            for (int i = 0; i < infoArr.length - 1; i++) {//Gets rid of all entries not inbetween the start and end date
                question = dateFormat.parse(infoArr[i][0]);
                if (question.compareTo(start) >= 0 && question.compareTo(end) <= 0) {
                    //count++;
                    Entry tempEntry = new Entry();//creates an temporary entry object
                    tempEntry.setDate(infoArr[i][0]);
                    tempEntry.setType(infoArr[i][1]);
                    tempEntry.setTime(infoArr[i][2]);
                    tempEntry.setSubject(infoArr[i][3]);
                    tempEntry.setCampus(infoArr[i][4]);
                    tempEntry.setName(infoArr[i][5]);

                    typeArr.add(infoArr[i][1]);
                    typeCamp.add(infoArr[i][4]);
                    typeSubject.add(infoArr[i][3]);
                    searchRes.add(tempEntry);//entry object added to searchRes for later use
                }
            }
            typeArr = removeDuplicates(typeArr);
            typeCamp = removeDuplicates(typeCamp);
            typeSubject = removeDuplicates(typeSubject);
            System.out.println(typeArr.size());
            System.out.println(typeCamp.size());
            System.out.println(typeSubject.size());
            int placeHolder = 0;
            if (searchRes.size() > 0) {//filters searchRes by type of visit, campus, and subject
                if (query.equals("111111111111")) { //checks if the query is searching for everything
//                    placeHolder = 0;
//                    do {
//                        if (checkOdd(searchRes.get(placeHolder).getType()) == true) {//Checks if visit type is anything but Appointments or dropins and deletes them
//                            searchRes.remove(placeHolder);
//                        } else {
//                            placeHolder++;
//                        }
//                    } while (placeHolder < searchRes.size());
                }
                for (int i = 0; i < query.length(); i++) {//goes through each filter to remove from searchRes
                    placeHolder = 0;
                    do {
                        if (searchType(searchRes.get(placeHolder).getSubject(), "Citation/Reference") == true) {//removes any citation/reference entries as we do not report on them
                            searchRes.remove(placeHolder);
                        } else {
                            placeHolder++;
                        }
                    } while (placeHolder < searchRes.size());
                    switch (i) {
                        case 0://Appointments
                            placeHolder = 0;
                            if (query.charAt(i) == '0') { //removes any appointment entries 
                                do {
                                    if (searchType(searchRes.get(placeHolder).getType(), "Appointment") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 1://Drop ins
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any drop in entries
                                do {
                                    if (searchType(searchRes.get(placeHolder).getType(), "Drop-In") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 2: //Emails
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any drop in entries
                                do {
                                    if (searchType(searchRes.get(placeHolder).getType(), "Email") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 3://Davis
                            placeHolder = 0;
                            if (query.charAt(i) == '0') { //removes any entries from the davis campus
                                do {
                                    if (searchType(searchRes.get(placeHolder).getCampus(), "Davis") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 4://TRA
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any entries from the TRA campus
                                do {
                                    if (searchType(searchRes.get(placeHolder).getCampus(), "Trafalgar") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 5://HMC
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any entries from the HMC campus
                                do {
                                    if (searchType(searchRes.get(placeHolder).getCampus(), "HMC") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 6:
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any entries from the HMC campus
                                do {
                                    if (searchType(searchRes.get(placeHolder).getCampus(), "Online") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 7://Comp Sci
                            placeHolder = 0;
                            if (query.charAt(i) == '0') { //removes any entries of computer science/ java
                                do {
                                    if (searchType(searchRes.get(placeHolder).getSubject(), "Computer Programming") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 8://English
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any entries of english
                                do {
                                    if (searchType(searchRes.get(placeHolder).getSubject(), "English") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 9://Math
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any entries of math
                                do {
                                    if (searchType(searchRes.get(placeHolder).getSubject(), "Math") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 10://Business Math
                            placeHolder = 0;
                            if (query.charAt(i) == '0') { //removes any entries of business math
                                do {
                                    if (searchType(searchRes.get(placeHolder).getSubject(), "Business Math") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                        case 11://Online English
                            placeHolder = 0;
                            if (query.charAt(i) == '0') {//removes any entries of online english
                                do {
                                    if (searchType(searchRes.get(placeHolder).getSubject(), "Online English") == true) {
                                        searchRes.remove(placeHolder);
                                    } else {
                                        placeHolder++;
                                    }
                                } while (placeHolder < searchRes.size());
                            }
                            break;
                    }
                }
                startList.clear();
                for (int x = 0; x < searchRes.size(); x++) {
                    startList.add(searchRes.get(x).getName());
                }
                Set<String> endList = new HashSet<String>(startList);
                studentCount = endList.size();
            }
            outputWindow();
        } catch (ParseException ex) {
            Logger.getLogger(ExcelTouch.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//getInfo2

    public void outputWindow2(ArrayList<String> apps, ArrayList<String> camp, ArrayList<String> sub) {
        int[] timeTotals = new int[apps.size()];
        int[] numApps = new int[apps.size()];
        for (int i = 0; i < apps.size(); i++) {
            for (int j = 0; j < searchRes.size(); j++) {
                if (searchRes.get(j).getType().equals(apps.get(i))) {
                    timeTotals[i] += Integer.parseInt(searchRes.get(j).getTime());
                    numApps[i] += 1;
                }
            }
        }
        output = "These results are for ";
        //System.out.print("These results are for ");
        for (int i = 0; i < apps.size(); i++) {
            if (i == apps.size() - 1) {
                if (apps.size() > 1) {
                    output += "& ";
                }
                output += apps.get(i) + "s\n";
            } else {
                output += apps.get(i) + "s, ";
            }
        }
        output += "at ";
        for (int i = 0; i < camp.size(); i++) {
            if (i == camp.size() - 1) {
                if (camp.size() > 1) {
                    output += "& ";
                }
                output += camp.get(i) + "\n";
            } else {
                output += camp.get(i) + ", ";
            }
        }
        output += "for ";
        for (int i = 0; i < sub.size(); i++) {
            if (i == sub.size() - 1) {
                if (sub.size() > 1) {
                    output += "& ";
                }
                output += sub.get(i) + "\n";
            } else {
                output += sub.get(i) + ", ";
            }
        }
        output += "\n\n";
        for (int i = 0; i < apps.size(); i++) {
            output += "Number of " + apps.get(i) + "s :" + numApps[i] + "\n";
            output += "Total time for " + apps.get(i) + ": " + (timeTotals[i] / 60) + " hours and " + (timeTotals[i] % 60 + " minutes.\n\n");
        }
        output += "\nNumber of Distinct Students: " + studentCount;
    }

    public void outputWindow() { //creates an easy to read break down of number of appointments and drop-ins and the hours for each
        int numAppointments = 0;
        int numDropIns = 0;

        int numEmails = 0;
        double numMinutesAppointments = 0;
        double numMinutesDropIns = 0;
        double numMinutesEmails = 0;
        double hoursApp = 0;
        double hoursDrop = 0;
        double hoursEmail = 0;
        String typeString = "";
        String campusString = "";
        String subjectString = "";

        String formattedApp = "0";
        String formattedDrops = "0";
        String formattedEmails = "0";
        for (int i = 0; i < searchRes.size(); i++) {
            if (searchRes.get(i).getType().equals("Appointment")) {
                numAppointments++;
                numMinutesAppointments += Integer.parseInt(searchRes.get(i).getTime());
            } else if (searchRes.get(i).getType().equals("Drop-In")) {
                numDropIns++;
                numMinutesDropIns += Integer.parseInt(searchRes.get(i).getTime());
            } else if (searchRes.get(i).getType().equals("Email")) {
                numEmails++;
                numMinutesEmails += Integer.parseInt(searchRes.get(i).getTime());
            }
        }
        if (numMinutesAppointments > 0) {
            hoursApp = numMinutesAppointments / 60; // modulo on minutes
        }
        if (numMinutesDropIns > 0) {
            hoursDrop = numMinutesDropIns / 60;
        }
        if (numMinutesEmails > 0) {
            hoursEmail = numMinutesEmails / 60;
        }
        for (int i = 0; i < this.queryIn.length(); i++) {
            switch (i) {
                case 0:
                    if (queryIn.charAt(i) == '1') {
                        typeString += "Appointments";
                    }
                    break;
                case 1:
                    if (queryIn.charAt(i) == '1') {
                        if (typeString.equals("Appointments")) {
                            typeString += " and Drop-Ins";
                        } else {
                            typeString += "Drop-Ins";
                        }
                    }
                    break;
                case 2:
                    if (queryIn.charAt(i) == '1') {
                        if (typeString.equals("Appointments")) {
                            typeString += " and Drop-Ins";
                            typeString += " and Emails";
                        } else {
                            typeString += " Emails";
                        }
                    }
                    break;
                case 3:
                    if (queryIn.charAt(i) == '1') {
                        campusString += "Davis";
                    }
                    break;
                case 4:
                    if (queryIn.charAt(i) == '1') {
                        if (campusString.equals("Davis") && queryIn.charAt(i + 1) == '1') {
                            campusString += ", TRA,";
                        } else if (campusString.equals("Davis")) {
                            campusString += " and TRA";
                        } else {
                            campusString += "TRA";
                        }
                    }
                    break;
                case 5:
                    if (queryIn.charAt(i) == '1') {
                        if (queryIn.charAt(i - 2) == '1' ^ queryIn.charAt(i - 1) == '1' ^ (queryIn.charAt(i - 1) == '1' && queryIn.charAt(i - 2) == '1')) {
                            campusString += " and HMC";
                        } else {
                            campusString += "HMC";
                        }
                    }
                    break;
                case 6:
                    if (queryIn.charAt(i) == '1') {
                        if (queryIn.charAt(i - 2) == '1' ^ queryIn.charAt(i - 1) == '1' ^ (queryIn.charAt(i - 1) == '1' && queryIn.charAt(i - 2) == '1')) {
                            campusString += " and Online";
                        } else {
                            campusString += "Online campus";
                        }
                    }
                    break;
                case 7:
                    if (queryIn.charAt(i) == '1') {
                        subjectString += "Computer Science,";
                    }
                    break;
                case 8:
                    if (queryIn.charAt(i) == '1') {
                        if (queryIn.charAt(i - 1) == '1' && (queryIn.charAt(i + 1) == '1' || queryIn.charAt(i + 2) == '1' || queryIn.charAt(i + 3) == '1')) {
                            subjectString += " English,";
                        } else if (queryIn.charAt(i - 1) == '1') {
                            subjectString += " and English";
                        } else {
                            subjectString += "English,";
                        }
                    }
                    break;
                case 9:
                    if (queryIn.charAt(i) == '1') {
                        if ((queryIn.charAt(i - 1) == '1' || queryIn.charAt(i - 2) == '1') && (queryIn.charAt(i + 1) == '1' || queryIn.charAt(i + 2) == '1')) {
                            subjectString += " Math,";
                        } else if (queryIn.charAt(i - 1) == '1' || queryIn.charAt(i - 2) == '1') {
                            subjectString += " and Math";
                        } else {
                            subjectString += "Math,";
                        }
                    }
                    break;
                case 10:
                    if (queryIn.charAt(i) == '1') {
                        if ((queryIn.charAt(i - 1) == '1' || queryIn.charAt(i - 2) == '1' || queryIn.charAt(i - 3) == '1') && queryIn.charAt(i + 1) == '1') {
                            subjectString += " Business Math,";
                        } else if (queryIn.charAt(i - 1) == '1' || queryIn.charAt(i - 2) == '1' || queryIn.charAt(i - 3) == '1') {
                            subjectString += " and Business Math";
                        } else {
                            subjectString += "Business Math,";
                        }
                    }
                    break;
                case 11:
                    if (queryIn.charAt(i) == '1') {
                        if (queryIn.charAt(i - 1) == '1' || queryIn.charAt(i - 2) == '1' || queryIn.charAt(i - 3) == '1' || queryIn.charAt(i - 4) == '1') {
                            subjectString += " and Online English";
                        } else {
                            subjectString += "Online English";
                        }
                    }
                    break;
            }
        }

        DecimalFormat df2 = new DecimalFormat(".##");
        if (numAppointments > 0) {
            formattedApp = df2.format(numMinutesAppointments / numAppointments);
        }
        if (numDropIns > 0) {
            formattedDrops = df2.format(numMinutesDropIns / numDropIns);
        }
        if (numEmails > 0) {
            System.out.println("NUMBER OF EMAILS" + numEmails);
            formattedEmails = df2.format(numMinutesEmails / numEmails);
        }

        output = "Results for " + typeString + "\nat " + campusString + "\nfor " + subjectString + "\n\n";
        output += "Number of distinct Students: " + studentCount + "\n\n";
        if (queryIn.charAt(0) == '1') {
            output += "Number of Appointments: " + numAppointments + "\nHours of Appointments: " + df2.format(hoursApp) + "\n\n";
        }
        if (queryIn.charAt(1) == '1') {
            output += "Number of Drop-Ins: " + numDropIns + "\nHours of Drop-Ins: " + df2.format(hoursDrop) + "\n\n";
        }
        if (queryIn.charAt(2) == '1') {
            output += "Number of Emails: " + numEmails + "\nHours of Emails: " + df2.format(hoursEmail) + "\n\n";
        }
        output += "Total Visits: " + (numAppointments + numDropIns) + "\n\n";
        if (queryIn.charAt(0) == '1') {
            output += "That\'s " + formattedApp + " minutes per Appointment\n";
        }
        if (queryIn.charAt(1) == '1') {
            output += "That\'s " + formattedDrops + " minutes per Drop-In\n";
        }
        if (queryIn.charAt(2) == '1') {
            output += "That\'s " + formattedEmails + " minutes per Email Appointment\n";
        }
    }//outputWindow

    public boolean checkOdd(String compareThis) {//checks for any anomalies in visit type
        boolean matches = false;
        if (compareThis.equals("Email")) {
            matches = true;
        }
        return matches;
    }

    public boolean searchType(String compareThis, String toThis) {//compares two things as it took up too much space if i wrote it out for every single time i needed to check something
        boolean matches = false;
        if (compareThis.equals(toThis)) {
            matches = true;
        }
        return matches;
    }

    static ArrayList<String> removeDuplicates(ArrayList<String> list) {

        // Store unique items in result.
        ArrayList<String> result = new ArrayList<>();

        // Record encountered Strings in HashSet.
        HashSet<String> set = new HashSet<>();

        // Loop over argument list.
        for (String item : list) {

            // If String is not in set, add it to the list and the set.
            if (!set.contains(item)) {
                result.add(item);
                set.add(item);
            }
        }
        return result;
    }
} // ExcelTouch
