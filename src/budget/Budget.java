/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package budget;


import java.awt.MenuItem;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.GroupLayout;
import jxl.*;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 *
 * @author Bill
 */
public class Budget {
    //static String path = System.getProperty("user.dir");
    static WritableWorkbook workbookEdit;
    static WritableSheet sheetEdit, sheetEdit1, sheetEdit2;
    static Workbook workbook;
    static Sheet sheet, sheet1, sheet2;
    static Cell cellContents, prevDebitCell, prevCreditCell;
    static Formula formula;
    static NumberCell numberCell;
    static Scanner in = new Scanner(System.in);
    static boolean success;
    static String confirm, transNumString, transDesc, transCat, transDebitString, transCreditString, csvError, month, transType, transType2, transCatString, 
            transCatString2, initalBalanceString, accountTypeString, fileName, transDateString;
    static int rowCounter, columnCounter, transNumInt, transIndex, transIndex2, assetsLength, revenueLength, expensesLenth,sheetNum = 0,
            //summary variables
            accountTypeSummaryCol = 0, accountSummaryCol = 1, debitSummaryCol = 2, creditSummaryCol = 3, subheadingSummaryRow = 1, headingSummaryRow =0, headingSummaryCol = 0,
            //journal variables
            transNumJournalCol = 0, dateJournalCol = 1, descJournalCol = 2, accountJournalCol = 3, debitJournalCol =4, creditJournalCol = 5, subheadingJournalRow = 1, headingJournalRow = 0, headingJournalCol = 0,
            //T-Account variables
            headingTRow = 0, headingTCol = 0, assetTRow = 1, revenueTRow = 6, expensesTRow = 11;
    static double prevDebit, prevCredit, initialBalanceInt, assetDebitSummary, expenseDebitSummary, revenueCreditSummary
            , expenseCreditSummary, assetDebit, assetCredit, revenueDebit, revenueCredit, expenseDebit, expenseCredit, assetCreditSummary, revenueDebitSummary;
    static String[][] viewSummary, viewJournal, viewTAccounts;
    static DateFormat dateFormatMonth = new SimpleDateFormat("MM");
    static DateFormat dateFormatYear = new SimpleDateFormat("yyyy");
    static DateFormat dateFormatTrans = new SimpleDateFormat("MM/dd/yyyy");
    static Date date = new Date(), transDate = new Date();
    static WritableCellFormat journalFormatHeading, journalFormatHeading2, tAccountFormatHeading3, tAccountFormatDivider; 
    static WritableFont journalFontHeading, journalFontHeading2, tAccountFontHeading3, tAccountFontDivider;
    static ArrayList assetAccountsAL, revenueAccountsAL, expenseAccountsAL, assetInitialAL, revenueInitialAL, expenseInitialAL;
    static String[] accountTypes = {"Assets", "Revenue", "Expenses"};
    static Object[] assetAccounts, revenueAccounts, expenseAccounts, accountsObj, assetInitial, revenueInitial, expenseInitial;
    static String[] accounts;
    static char[] alphabet = {'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z'};
    static String dirFile = "..\\excel\\", tempDir = "temp\\";
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        
        journalFontHeading = new WritableFont(WritableFont.ARIAL, 14);
        journalFontHeading2 = new WritableFont(WritableFont.ARIAL, 10);
        tAccountFontHeading3 = new WritableFont(WritableFont.ARIAL, 10);
        tAccountFontDivider = new WritableFont(WritableFont.ARIAL, 10);
        try {
            journalFontHeading.setBoldStyle(WritableFont.BOLD);
            journalFontHeading2.setBoldStyle(WritableFont.BOLD);
            tAccountFontHeading3.setBoldStyle(WritableFont.BOLD);
        } catch (WriteException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
        journalFormatHeading = new WritableCellFormat(new WritableFont(journalFontHeading));
        journalFormatHeading2 = new WritableCellFormat(new WritableFont(journalFontHeading2));
        tAccountFormatHeading3 = new WritableCellFormat(new WritableFont(tAccountFontHeading3));
        tAccountFormatDivider = new WritableCellFormat(new WritableFont(tAccountFontDivider));
        try {
            journalFormatHeading.setAlignment(Alignment.CENTRE);
            journalFormatHeading2.setAlignment(Alignment.CENTRE);
            tAccountFormatHeading3.setAlignment(Alignment.CENTRE);
            tAccountFormatHeading3.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
            tAccountFormatDivider.setBorder(Border.LEFT, BorderLineStyle.THIN);
        } catch (WriteException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
        assetAccountsAL = new ArrayList();
        revenueAccountsAL = new ArrayList();
        expenseAccountsAL = new ArrayList();
        assetInitialAL = new ArrayList();
        revenueInitialAL = new ArrayList();
        expenseInitialAL = new ArrayList();
        switch(dateFormatMonth.format(date)){
            case "01":
                month = "January";
                break;
            case "02":
                month = "February";
                break;
            case "03":
                month = "March";
                break;
            case "04":
                month = "April";
                break;
            case "05":
                month = "May";
                break;           
            case"06":
                month = "June";
                break;
            case "07":
                month = "July";
                break;
            case "08":
                month = "August";
                break;
            case "09":
                month = "September";
                break;
            case "10":
                month = "October";
                break;
            case "11":
                month = "November";
                break;
            
            default:
                month = "error";
                break;

        }
        new welcome().setVisible(true);
        
    }
    /*
     * a function to create a new CSV file
     */
    public static void createCSV(boolean initial) throws WriteException{
        
        assetAccounts = assetAccountsAL.toArray();
        revenueAccounts = revenueAccountsAL.toArray();
        expenseAccounts = expenseAccountsAL.toArray();
        assetInitial = assetInitialAL.toArray();
        revenueInitial = revenueInitialAL.toArray();
        expenseInitial = expenseInitialAL.toArray();
        
        //create workbook
        try
	{
            //accountIndex = new int[accounts.length];
	    workbookEdit = Workbook.createWorkbook(new File(dirFile + fileName + ".xls"));
            sheetEdit = workbookEdit.createSheet(month + " " + dateFormatYear.format(date) + " Summary", 0);
            sheetEdit1 = workbookEdit.createSheet(month + " " + dateFormatYear.format(date) + " Journal", 1);
            sheetEdit2 = workbookEdit.createSheet(month + " " + dateFormatYear.format(date) + " T-Accounts", 2);
            try { 
                sheetEdit.mergeCells(0, 0, 3, 0);
                sheetEdit.addCell(new Label(headingSummaryCol, headingSummaryRow, "Summary", journalFormatHeading));
                sheetEdit.addCell(new Label(accountTypeSummaryCol, subheadingSummaryRow, "Account Type", journalFormatHeading2));
                sheetEdit.addCell(new Label(accountSummaryCol, subheadingSummaryRow, "Account", journalFormatHeading2));
                sheetEdit.addCell(new Label(debitSummaryCol, subheadingSummaryRow, "Debit", journalFormatHeading2));
                sheetEdit.addCell(new Label(creditSummaryCol, subheadingSummaryRow, "Credit", journalFormatHeading2));
                //building accounts Array
                accountsObj = new Object[assetAccounts.length + revenueAccounts.length + expenseAccounts.length];
                for(int i = 0; i < assetAccounts.length; i++){
                    accountsObj[i] = assetAccounts[i];
                    sheetEdit.addCell(new Label(accountTypeSummaryCol, i+2, accountTypes[0]));
                    sheetEdit.addCell(new Label(accountSummaryCol, i+2, assetAccounts[i].toString()));
                    
                }
                for(int i = 0; i < revenueAccounts.length; i++){
                    accountsObj[i + assetAccounts.length] = revenueAccounts[i];
                    sheetEdit.addCell(new Label(accountTypeSummaryCol, i+2+assetAccounts.length, accountTypes[1]));
                    sheetEdit.addCell(new Label(accountSummaryCol, i+2+assetAccounts.length, revenueAccounts[i].toString()));
                    
                }
                for(int i = 0; i < expenseAccounts.length; i++){
                    accountsObj[i + (assetAccounts.length + revenueAccounts.length)] = expenseAccounts[i];
                    sheetEdit.addCell(new Label(accountTypeSummaryCol, i+2+ (assetAccounts.length + revenueAccounts.length), accountTypes[2]));
                    sheetEdit.addCell(new Label(accountSummaryCol, i+2+ (assetAccounts.length + revenueAccounts.length), expenseAccounts[i].toString()));
                    
                }
                /*
                 * Journal
                 */    
                
                //transaction headers
                sheetEdit1.mergeCells(0, 0, 4, 0);
                //center this
                sheetEdit1.addCell(new Label(headingJournalCol, headingJournalRow, "Journal", journalFormatHeading));
                sheetEdit1.addCell(new Label(transNumJournalCol, subheadingJournalRow, "Transaction Number", journalFormatHeading2));
                sheetEdit1.addCell(new Label(dateJournalCol, subheadingJournalRow, "Transaction Date", journalFormatHeading2));
                sheetEdit1.addCell(new Label(descJournalCol, subheadingJournalRow, "Description", journalFormatHeading2));
                sheetEdit1.addCell(new Label(accountJournalCol, subheadingJournalRow, "Account", journalFormatHeading2));
                sheetEdit1.addCell(new Label(debitJournalCol, subheadingJournalRow, "Debit", journalFormatHeading2));
                sheetEdit1.addCell(new Label(creditJournalCol, subheadingJournalRow, "Credit", journalFormatHeading2));
                /*
                 * T-Accounts
                 */
                //T-account headers
                sheetEdit2.mergeCells(0, 0, 20, 0);
                sheetEdit2.addCell(new Label(headingTCol, headingTRow, "T-Accounts", journalFormatHeading));
                
                sheetEdit2.mergeCells(0, 1, 20, 1);
                sheetEdit2.addCell(new Label(headingTCol, assetTRow, "Assets", journalFormatHeading2));
                
                
                for(int i = 0; i < assetAccounts.length; i++){
                    sheetEdit2.mergeCells(i*4, 2, i*4+3, 2);
                   sheetEdit2.addCell(new Label(i * 4, 2, assetAccounts[i].toString(), tAccountFormatHeading3));
                   sheetEdit2.addCell(new Label(i*4, 3, "Debit"));
                   sheetEdit2.addCell(new Label(i*4+2, 3, "Credit", tAccountFormatDivider));
                   sheetEdit2.mergeCells(i*4, 4, i*4+1, 4);
                   sheetEdit2.mergeCells(i*4+2, 4, i*4+3, 4);
                   sheetEdit2.addCell(new Label(i*4, 4, "Total"));
                   
                   //formula = new Formula((i*4)+2, 4, Character.toString(alphabet[(4*i)+1]) + "4-" + Character.toString(alphabet[(4*i)+3]) + "4");
                    //sheetEdit2.addCell(formula);
                }
                sheetEdit2.mergeCells(0, 6, 20, 6);
                sheetEdit2.addCell(new Label(headingTCol, revenueTRow, "Revenue", journalFormatHeading2));
                for(int i = 0; i < revenueAccounts.length; i++){
                    sheetEdit2.mergeCells(i*4, 7, i*4+3, 7);
                   sheetEdit2.addCell(new Label(i * 4, 7, revenueAccounts[i].toString(), tAccountFormatHeading3));
                   sheetEdit2.addCell(new Label(i*4, 8, "Debit"));
                   sheetEdit2.addCell(new Label(i*4+2, 8, "Credit", tAccountFormatDivider));
                   sheetEdit2.mergeCells(i*4, 9, i*4+1, 9);
                   sheetEdit2.mergeCells(i*4+2, 9, i*4+3, 9);
                   sheetEdit2.addCell(new Label(i*4, 9, "Total"));
                   //formula = new Formula((i*4)+2, 9, Character.toString(alphabet[(4*i)+3]) + "9-" + Character.toString(alphabet[(4*i)+1]) + "9");
                    //sheetEdit2.addCell(formula);
                }
                sheetEdit2.mergeCells(0, 11, 20, 11);
                sheetEdit2.addCell(new Label(headingTCol, expensesTRow, "Expenses", journalFormatHeading2));
                for(int i = 0; i < expenseAccounts.length; i++){
                    sheetEdit2.mergeCells(i*4, 12, i*4+3, 12);
                   sheetEdit2.addCell(new Label(i * 4, 12, expenseAccounts[i].toString(), tAccountFormatHeading3));
                   sheetEdit2.addCell(new Label(i*4, 13, "Debit"));
                   sheetEdit2.addCell(new Label(i*4+2, 13, "Credit", tAccountFormatDivider));
                   sheetEdit2.mergeCells(i*4, 14, i*4+1, 14);
                   sheetEdit2.mergeCells(i*4+2, 14, i*4+3, 14);
                   sheetEdit2.addCell(new Label(i*4, 14, "Total"));
                   /*
                    * Initial Balance
                    */
                   
                       
                            
                            
                   
                   //formula = new Formula((i*4)+2, 14, Character.toString(alphabet[(4*i)+1]) + "14-" + Character.toString(alphabet[(4*i)+3]) + "14");
                    //sheetEdit2.addCell(formula);
                }
                //for(int i = 0; i < assetAccounts.length; i++){
                    //formula = new Formula(2, i+2,"'june 2014 T-Accounts'!" + Character.toString(alphabet[(4*i)+2]) + "5");
                    //sheetEdit.addCell(formula);
                //}for(int i = 0; i < revenueAccounts.length; i++){
                    //formula = new Formula(3, i+2+assetAccounts.length,"'june 2014 T-Accounts'!" + Character.toString(alphabet[(4*i)+2]) + "10");
                    //sheetEdit.addCell(formula);
                //}for(int i = 0; i < expenseAccounts.length; i++){
                    //formula = new Formula(2, i+2+ (assetAccounts.length + revenueAccounts.length),"'june 2014 T-Accounts'!" + Character.toString(alphabet[(4*i)+2]) + "15");
                    //sheetEdit.addCell(formula);
                //}
                
                
                
                //sheetEdit.addCell(new Label(6, 1, ))
            //sheetNew.addCell(new Number(3, 4, 3.1459));
                workbookEdit.write(); 
                workbookEdit.close();
            } catch (WriteException ex) {
                Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
            }
        for(int i = 0; i < assetAccounts.length; i++){
            if(assetInitial[i]!=0){
                Budget.enterInitialTransaction(assetAccounts[i].toString(), "Assets", (Double)assetInitial[i]);
            }
        }
        for(int i = 0; i < revenueAccounts.length; i++){
            if(assetInitial[i]!=0){
                Budget.enterInitialTransaction(revenueAccounts[i].toString(), "Revenue", (Double)revenueInitial[i]);
            }
        }
        for(int i = 0; i < expenseAccounts.length; i++){
            if(assetInitial[i]!=0){
                Budget.enterInitialTransaction(expenseAccounts[i].toString(), "Expenses", (Double)expenseInitial[i]);
            }
        }

            
        }
	catch(IOException e)
	{
	     e.printStackTrace();
	} 
        
        
    }
    /*
     * funtion to open existing workbook
     */
    public static void openExisting(String fileName){
        Budget.fileName = fileName;    
        new loadFile().setVisible(false);
            update();
        
    }
    /*
     * a function to update an existing CSV file
    */ 
    public static void update(){
        try {

            workbook = Workbook.getWorkbook(new File(dirFile + fileName + ".xls"));
            sheet = workbook.getSheet(sheetNum);
            sheet1 = workbook.getSheet(sheetNum+1);
            sheet2 = workbook.getSheet(sheetNum+2);
            workbookEdit = Workbook.createWorkbook(new File(dirFile + tempDir + fileName + "_temp.xls"), workbook);
            sheetEdit = workbookEdit.getSheet(sheetNum);
            sheetEdit1 = workbookEdit.getSheet(sheetNum+1);
            sheetEdit2 = workbookEdit.getSheet(sheetNum+2);
            transNumGet();
                
            
            transNumString = "" + transNumInt;
            //new loading().setVisible(true);
            saveArrays();
 
            csvCommand();
        } catch (BiffException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }catch (IOException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    /*
     * function to give you the main list of commands
    */ 
    public static void csvCommand(){
        
        new command().setVisible(true);
    }
    
     
     
    public static void viewContents(){
        viewSummary = new String[sheetEdit.getRows() - 2][sheetEdit.getColumns()];
        for(int i = 0; i<sheetEdit.getColumns();i++){
            for(int j = 0; j<sheetEdit.getRows() - 2;j++){
                cellContents = sheetEdit.getCell(i, j+2);
                
                viewSummary[j][i] = cellContents.getContents();
                //System.out.println(cellContents.getContents());
            }
        }
        viewJournal = new String[sheetEdit1.getRows() - 2][sheetEdit1.getColumns()];
        for(int i = 0; i<sheetEdit1.getColumns();i++){
            for(int j = 0; j<sheetEdit1.getRows() - 2;j++){
                cellContents = sheetEdit1.getCell(i, j+2);
                viewJournal[j][i] = cellContents.getContents();
            }
        }
       
        new view(viewSummary, viewJournal).setVisible(true);
        //csvCommand();
    }
    /*
     * function to add accounts
     */
    public static void addAccounts(String accountName, Object accountTypeObj, String initialBalanceString){
        
        if(initialBalanceString.isEmpty()==false){
            initialBalanceInt = Double.parseDouble(initialBalanceString);
        }else{
            initialBalanceInt = 0;
        }
        accountTypeString = accountTypeObj.toString();
        if(accountTypeString.matches(accountTypes[0])){
            assetAccountsAL.add(accountName);
            assetInitialAL.add(initialBalanceInt);
        }
        else if(accountTypeString.matches(accountTypes[1])){
            revenueAccountsAL.add(accountName);
            revenueInitialAL.add(initialBalanceInt);
        }
        else if(accountTypeString.matches(accountTypes[2])){
            expenseAccountsAL.add(accountName);
            expenseInitialAL.add(initialBalanceInt);
        }else{
            System.out.println("Error with adding account");
            
        }
        
    }
    public static void addMoreAccounts(){
        new newAccount(accountTypes).setVisible(true);
    }
    /*
     * function to add another new Account
     * may not be need
     */
    public static void addNewAccounts(String fileName){
        Budget.fileName = fileName;
        new newAccount(accountTypes).setVisible(true);
    }
    /*
     * function to get account names for use with the dropdown menu
     */
    public static String[] getAccounts(){
        
        return accounts;
    }
    /*
     * funtion to enter transactions into the spreadsheet
    */ 
    /*
     * function to repopulate arrays
     */
    public static void saveArrays(){
        accounts = new String[sheetEdit.getRows() - 2];
        for(int i = 0; i< accounts.length;i++){
            cellContents = sheetEdit.getCell(0, i+2);
            if(cellContents.getContents().matches(accountTypes[0])){
                assetsLength = i+1;
                
            }else if(cellContents.getContents().matches(accountTypes[1])){
                revenueLength = i+1;
                
            }else if(cellContents.getContents().matches(accountTypes[2])){
                expensesLenth = i+1;
                
            }
            
        }
        
            if(assetsLength!=0){
                assetAccounts = new String[assetsLength];
            }    
        
            
            if(revenueLength!=0){
                revenueLength -= assetsLength;
                revenueAccounts = new String[revenueLength];
            }
        
        
            
            if(expensesLenth!=0){
                expensesLenth -= (assetsLength + revenueLength);
                expenseAccounts = new String[expensesLenth];
            }
        

        for(int i = 0; i < accounts.length; i++){
            cellContents = sheetEdit.getCell(accountSummaryCol, i+2);
            accounts[i] = cellContents.getContents();
            cellContents = sheetEdit.getCell(accountTypeSummaryCol, i+2);
            //System.out.println(i);
            if(cellContents.getContents().matches(accountTypes[0])){
                cellContents = sheetEdit.getCell(accountSummaryCol, i+2);
                assetAccounts[i] = cellContents.getContents();
                cellContents = sheetEdit.getCell(accountTypeSummaryCol, i+2);
            }else if(cellContents.getContents().matches(accountTypes[1])){
                cellContents = sheetEdit.getCell(accountSummaryCol, i+2);
                 revenueAccounts[i - assetsLength] = cellContents.getContents();
                cellContents = sheetEdit.getCell(accountTypeSummaryCol, i+2);
            }else if(cellContents.getContents().matches(accountTypes[2])){
                cellContents = sheetEdit.getCell(accountSummaryCol, i+2);
                expenseAccounts[i - (assetsLength + revenueLength)] = cellContents.getContents();
                cellContents = sheetEdit.getCell(accountTypeSummaryCol, i+2);
            }else{
                System.out.println("Error: No Accounts Match");
            }
        }

    }
    /*
     * Method that handles writing the transaction to the spreadsheet
     */
    public static void enterTransaction(String transDate, String transDesc, Object transCat, String transDebitString, Object transCat2, String transCredit2String){
        try {
            transNumGet();
            transCatString = transCat.toString();
            transCatString2 = transCat2.toString();
                journalTransaction(transNumInt, transDate, transDesc, transCatString, transDebitString, transCatString2, transCredit2String);
                transIndices();
                TandSDebitTransaction(transDesc, transCatString, transDebitString, transCatString2, transCredit2String);
                saveWorkbook();
                saveArrays();
                loadWorkbook();
                TandSCreditTransaction(transDesc, transCatString, transDebitString, transCatString2, transCredit2String);
                saveWorkbook();
                update();
        } catch (WriteException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    public static void journalTransaction(int transNumInt, String transDate, String transDesc, String transCatString, String transDebitString, String transCatString2, String transCredit2String){
        try {
            
                    
                sheetEdit1.addCell(new Number(transNumJournalCol, rowCounter, transNumInt));
                sheetEdit1.addCell(new Label(dateJournalCol, rowCounter, transDate));
                sheetEdit1.addCell(new Label(descJournalCol, rowCounter, transDesc));
                sheetEdit1.addCell(new Label(accountJournalCol, rowCounter, transCatString));
                
                sheetEdit1.addCell(new Label(debitJournalCol, rowCounter, transDebitString));
                sheetEdit1.addCell(new Label(accountJournalCol, rowCounter+1, transCatString2));
                sheetEdit1.addCell(new Label(creditJournalCol, rowCounter+1, transCredit2String));
        } catch (WriteException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    /*
     * method to get indices of where to place cells
     */
    public static void transIndices(){
        for(int i = 0; i<assetsLength;i++){
                if(transCatString.matches(assetAccounts[i].toString())){
                    transType = accountTypes[0];
                    transIndex = i;
                }
                if(transCatString2.matches(assetAccounts[i].toString())){
                    transType2 = accountTypes[0];
                    transIndex2 = i;
                }
            }
            for(int i = 0; i<revenueLength;i++){
                if(transCatString.matches(revenueAccounts[i].toString())){
                    transType = accountTypes[1];
                    transIndex = i;
                }
                if(transCatString2.matches(revenueAccounts[i].toString())){
                    transType2 = accountTypes[1];
                    transIndex2 = i;
                }
            }
            for(int i = 0; i<expensesLenth;i++){
                if(transCatString.matches(expenseAccounts[i].toString())){
                    transType = accountTypes[2];
                    transIndex = i;
                }
                if(transCatString2.matches(expenseAccounts[i].toString())){
                    transType2 = accountTypes[2];
                    transIndex2 = i;
                }
            }
    }
    /*
     * Method to enter debit transactions for Summary and T-Accounts
     */
    public static void TandSDebitTransaction(String transDesc, String transCatString, String transDebitString, String transCatString2, String transCredit2String){
         //find proper cells for adding
            if(transType.matches(accountTypes[0])){
            try {
                if(!transDebitString.matches("0")){
                    prevDebitCell = sheet2.getCell(transIndex*4+1, 3);
                    if(prevDebitCell.getContents().isEmpty()){
                        prevDebit = 0;
                    }else{
                        prevDebit = Double.parseDouble(prevDebitCell.getContents());
                    }
                }else{
                    prevDebit = 0;
                }
                sheetEdit2.addCell(new Number(transIndex*4+1, 3, prevDebit + Double.parseDouble(transDebitString)));
                //get corresponding credit
                cellContents = sheetEdit2.getCell(transIndex*4+3, 3);
                if(cellContents.getContents().isEmpty()){
                    assetCredit = 0;
                }else{
                    assetCredit = Double.parseDouble(cellContents.getContents());
                }
                
                sheetEdit2.addCell(new Number(transIndex*4+2, 4, prevDebit + Double.parseDouble(transDebitString) - assetCredit));
                sheetEdit.addCell(new Number(debitSummaryCol, transIndex+2, prevDebit + Double.parseDouble(transDebitString) - assetCredit));
            } catch (WriteException ex) {
                Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
            }
            }
            if(transType.matches(accountTypes[1])){
            try {
                if(!transDebitString.matches("0")){
                    prevDebitCell = sheet2.getCell(transIndex*4+1, 8);
                    if(prevDebitCell.getContents().isEmpty()){
                        prevDebit = 0;
                    }else{
                        prevDebit = Double.parseDouble(prevDebitCell.getContents());
                    }
                }else{
                    prevDebit = 0;
                }
                sheetEdit2.addCell(new Number(transIndex*4+1, 8, prevDebit + Double.parseDouble(transDebitString)));
                
                //get corresponding credit
                cellContents = sheetEdit2.getCell(transIndex*4+3, 8);
                if(cellContents.getContents().isEmpty()){
                    revenueCredit = 0;
                }else{
                    revenueCredit = Double.parseDouble(cellContents.getContents());
                }
                sheetEdit2.addCell(new Number(transIndex*4+2, 9, revenueCredit - (prevDebit + Double.parseDouble(transDebitString))));
                sheetEdit.addCell(new Number(creditSummaryCol, transIndex+2+assetsLength, revenueCredit - (prevDebit + Double.parseDouble(transDebitString))));
            } catch (WriteException ex) {
                Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
            } 
                
            }
            if(transType.matches(accountTypes[2])){
            try {
                if(!transDebitString.matches("0")){
                    prevDebitCell = sheet2.getCell(transIndex*4+1, 13);
                    if(prevDebitCell.getContents().isEmpty()){
                        prevDebit = 0;
                    }else{
                        prevDebit = Double.parseDouble(prevDebitCell.getContents());
                    }
                }else{
                    prevDebit = 0;
                }
                sheetEdit2.addCell(new Number(transIndex*4+1, 13, prevDebit + Double.parseDouble(transDebitString)));
                //sheetEdit.addCell(new Number(debitSummaryCol, transIndex+2+assetsLength+revenueLength, prevDebit + Double.parseDouble(transDebitString)));
                cellContents = sheetEdit2.getCell(transIndex*4+3, 13);
                if(cellContents.getContents().isEmpty()){
                    expenseCredit = 0;
                }else{
                    expenseCredit = Double.parseDouble(cellContents.getContents());
                }
                
                sheetEdit2.addCell(new Number(transIndex*4+2, 14, prevDebit + Double.parseDouble(transDebitString) - expenseCredit));
                sheetEdit.addCell(new Number(debitSummaryCol, transIndex+2+assetsLength+revenueLength, prevDebit + Double.parseDouble(transDebitString) - expenseCredit));
            } catch (WriteException ex) {
                Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
            }
            }
    }
    /*
     * Method for credit transactions for Summary and T-accounts
     */
    public static void TandSCreditTransaction(String transDesc, String transCatString, String transDebitString, String transCatString2, String transCredit2String){
        //Credits
            if(transType2.matches(accountTypes[0])){
            try {
                if(!transCredit2String.matches("0")){
                    prevCreditCell = sheet2.getCell(transIndex2*4+3, 3);
                    if(prevCreditCell.getContents().isEmpty()){
                        prevCredit = 0;
                    }else{
                        prevCredit = Double.parseDouble(prevCreditCell.getContents());
                    }
                }else{
                    prevCredit = 0;
                }
                sheetEdit2.addCell(new Number(transIndex2*4+3, 3, prevCredit + Double.parseDouble(transCredit2String)));
                //sheetEdit.addCell(new Number(debitSummaryCol, transIndex2+2, prevCredit + Double.parseDouble(transCredit2String)));
                //get corresponding debit
                cellContents = sheetEdit2.getCell(transIndex2*4+1, 3);
                if(cellContents.getContents().isEmpty()){
                    assetDebit = 0;
                }else{
                    assetDebit = Double.parseDouble(cellContents.getContents());
                }
                
                
                sheetEdit2.addCell(new Number(transIndex2*4+2, 4, assetDebit - (prevCredit + Double.parseDouble(transCredit2String))));
                sheetEdit.addCell(new Number(debitSummaryCol, transIndex2+2, assetDebit - (Double.parseDouble(transCredit2String) + prevCredit)));
            } catch (WriteException ex) {
                Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
            }
            }
            if(transType2.matches(accountTypes[1])){
            try {
                if(!transCredit2String.matches("0")){
                    prevCreditCell = sheet2.getCell(transIndex2*4+3, 8);
                    if(prevCreditCell.getContents().isEmpty()){
                        prevCredit = 0;
                    }else{
                        prevCredit = Double.parseDouble(prevCreditCell.getContents());
                    }
                }else{
                    prevCredit = 0;
                }
                sheetEdit2.addCell(new Number(transIndex2*4+3, 8, prevCredit + Double.parseDouble(transCredit2String)));
                //sheetEdit.addCell(new Number(creditSummaryCol, transIndex2+2+assetsLength, prevCredit + Double.parseDouble(transCredit2String)));
                //get corresponding debit
                cellContents = sheetEdit2.getCell(transIndex2*4+1, 8);
                if(cellContents.getContents().isEmpty()){
                    revenueDebit = 0;
                }else{
                    revenueDebit = Double.parseDouble(cellContents.getContents());
                }
                sheetEdit2.addCell(new Number(transIndex2*4+2, 9, prevCredit + Double.parseDouble(transCredit2String) - revenueDebit));
                sheetEdit.addCell(new Number(creditSummaryCol, transIndex2+2+assetsLength, prevCredit + Double.parseDouble(transCredit2String) - revenueDebitSummary));
            } catch (WriteException ex) {
                Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
            }
            }
            if(transType2.matches(accountTypes[2])){
            try {
                if(!transCredit2String.matches("0")){
                    prevCreditCell = sheet2.getCell(transIndex2*4+3, 13);
                    if(prevCreditCell.getContents().isEmpty()){
                        prevCredit = 0;
                    }else{
                        prevCredit = Double.parseDouble(prevCreditCell.getContents());
                    }
                }else{
                    prevCredit = 0;
                }
                sheetEdit2.addCell(new Number(transIndex2*4+3, 13, prevCredit + Double.parseDouble(transCredit2String)));
                //sheetEdit.addCell(new Number(debitSummaryCol, transIndex2+2+assetsLength+revenueLength, prevCredit + Double.parseDouble(transCredit2String)));
                //get corresponding debit
                cellContents = sheetEdit2.getCell(transIndex2*4+1, 13);
                if(cellContents.getContents().isEmpty()){
                    expenseDebit = 0;
                }else{
                    expenseDebit = Double.parseDouble(cellContents.getContents());
                }
                sheetEdit2.addCell(new Number(transIndex2*4+2, 14, expenseDebit - (prevCredit + Double.parseDouble(transCredit2String))));
                sheetEdit.addCell(new Number(debitSummaryCol, transIndex2+2+assetsLength+revenueLength, expenseDebit - (prevCredit + Double.parseDouble(transCredit2String))));
            } catch (WriteException ex) {
                Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
            }
            }
    }
    /*
     * Method for Saving Workbooks
     */
    public static void saveWorkbook() throws WriteException{
        try {
            workbookEdit.write(); 
               workbookEdit.close();
               workbook.close();
               workbook = Workbook.getWorkbook(new File(dirFile + tempDir + fileName + "_temp.xls"));
               workbookEdit = Workbook.createWorkbook(new File(dirFile + fileName + ".xls"), workbook);
               workbookEdit.write();
               workbookEdit.close();
               workbook.close();
               System.out.println("done");
        } catch (BiffException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    /*
     * Method that opens the workbooks
     */
    public static void loadWorkbook(){
        try {
            workbook = Workbook.getWorkbook(new File(dirFile + fileName + ".xls"));
                sheet = workbook.getSheet(sheetNum);
                sheet1 = workbook.getSheet(sheetNum+1);
                sheet2 = workbook.getSheet(sheetNum+2);
                workbookEdit = Workbook.createWorkbook(new File(dirFile + tempDir + fileName + "_temp.xls"), workbook);
                sheetEdit = workbookEdit.getSheet(sheetNum);
                sheetEdit1 = workbookEdit.getSheet(sheetNum+1);
                sheetEdit2 = workbookEdit.getSheet(sheetNum+2);
        } catch (IOException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        } catch (BiffException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    /*
     * Method that handles writing the transaction to the spreadsheet
     */
    public static void enterInitialTransaction(String accountName, String Type, double initalAccountBalanceDouble){
        try {
            workbook = Workbook.getWorkbook(new File(dirFile + fileName + ".xls"));
            sheet = workbook.getSheet(sheetNum);
            sheet1 = workbook.getSheet(sheetNum+1);
            sheet2 = workbook.getSheet(sheetNum+2);
            workbookEdit = Workbook.createWorkbook(new File(dirFile + tempDir + fileName + "_temp.xls"), workbook);
            sheetEdit = workbookEdit.getSheet(sheetNum);
            sheetEdit1 = workbookEdit.getSheet(sheetNum+1);
            sheetEdit2 = workbookEdit.getSheet(sheetNum+2);
            saveArrays();
            /*
             * Journal
             */
            
            transNumGet();
            
            
            //new Write().setVisible(true);
            transCatString = accountName;
            
            sheetEdit1.addCell(new Number(transNumJournalCol, rowCounter-1, transNumInt));
            sheetEdit1.addCell(new Label(dateJournalCol, rowCounter-1, dateFormatTrans.format(transDate)));
            sheetEdit1.addCell(new Label(descJournalCol, rowCounter-1, "Initialize"));
            sheetEdit1.addCell(new Label(accountJournalCol, rowCounter-1, transCatString));
            if(Type.matches(accountTypes[0].toString()) || Type.matches(accountTypes[2].toString())){
                sheetEdit1.addCell(new Number(debitJournalCol, rowCounter-1, initalAccountBalanceDouble));
            }
            if(Type.matches(accountTypes[1].toString())){
                sheetEdit1.addCell(new Number(creditJournalCol, rowCounter-1, initalAccountBalanceDouble));
            }
            /*
             * Summary and T-Accounts
             */
            if(Type.matches(accountTypes[0])){
                for(int i = 0; i<assetsLength;i++){
                    
                    if(transCatString.matches(assetAccounts[i].toString())){
                        transType = accountTypes[0];
                        transIndex = i;
                    }

                }
        }if(Type.matches(accountTypes[1])){
                for(int i = 0; i<revenueLength;i++){
                    if(transCatString.matches(revenueAccounts[i].toString())){
                        transType = accountTypes[1];
                        transIndex = i;
                    }

                }
        }if(Type.matches(accountTypes[2])){
                for(int i = 0; i<expensesLenth;i++){
                    if(transCatString.matches(expenseAccounts[i].toString())){
                        transType = accountTypes[2];
                        transIndex = i;
                    }

                }
        }
            //find proper cells for adding
            if(Type.matches(accountTypes[0])){
                
                sheetEdit2.addCell(new Number(transIndex*4+1, 3, initalAccountBalanceDouble));
                sheetEdit.addCell(new Number(debitSummaryCol, transIndex+2, initalAccountBalanceDouble));
                
                sheetEdit2.addCell(new Number(transIndex*4+2, 4, initalAccountBalanceDouble));
                //sheetEdit.addCell(new Number(2, transIndex+2, Double.parseDouble(transDebitString) - assetCreditSummary));
            }
            
            if(Type.matches(accountTypes[2])){
                sheetEdit2.addCell(new Number(transIndex*4+1, 13, prevDebit + initalAccountBalanceDouble));
                sheetEdit.addCell(new Number(debitSummaryCol, transIndex+2+assetsLength+revenueLength, prevDebit + initalAccountBalanceDouble));
                
                sheetEdit2.addCell(new Number(transIndex*4+2, 14, initalAccountBalanceDouble - expenseCredit));
                //sheetEdit.addCell(new Number(2, transIndex+2+assetsLength+revenueLength, Double.parseDouble(transDebitString) - expenseCreditSummary));
            }
            //Credits
            if(Type.matches(accountTypes[1])){
            
                
                sheetEdit2.addCell(new Number(transIndex*4+3, 8, prevCredit + initalAccountBalanceDouble));
                sheetEdit.addCell(new Number(creditSummaryCol, transIndex+2+assetsLength, prevCredit + initalAccountBalanceDouble));
                
                sheetEdit2.addCell(new Number(transIndex*4+2, 9, initalAccountBalanceDouble - revenueDebit));
                //sheetEdit.addCell(new Number(3, transIndex+2, Double.parseDouble(transCredit2String) - revenueDebitSummary));
            }
            
            
            
            
            workbookEdit.write(); 
            workbookEdit.close();
            workbook.close();
            workbook = Workbook.getWorkbook(new File(dirFile + tempDir + fileName + "_temp.xls"));
            workbookEdit = Workbook.createWorkbook(new File(dirFile + fileName + ".xls"), workbook);
            workbookEdit.write();
            workbookEdit.close();
            workbook.close();
            System.out.println("done");
        } catch (BiffException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WriteException ex) {
            Logger.getLogger(Budget.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    /*
     * Method to get the proper transaction Number
     */
    public static void transNumGet(){
        rowCounter = sheet1.getRows() + 1;
            transNumInt = rowCounter;
            transNumInt-=2;
            for(int i = 2; i < sheet1.getRows();i++){
                cellContents = sheet1.getCell(1, i);
                if(cellContents.getContents().isEmpty()){
                    transNumInt-=1;
                }
            }
    }
    /*
     * function to exit the program
     */
    public static void csvExit(){
        System.exit(0);
    }
}

