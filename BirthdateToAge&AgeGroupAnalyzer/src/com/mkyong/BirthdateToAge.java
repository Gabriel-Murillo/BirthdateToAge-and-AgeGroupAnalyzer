package com.mkyong;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//This is my third and final version. Most efficient, taking only 10 seconds. 
public class BirthdateToAge {
	private static final String FILE_NAME = "C:\\Users\\gabri\\Documents\\GitHub\\Excel Reader\\excelFiles\\VotingInfoDoc1.xlsx";
    private static ArrayList<String> bDList = new ArrayList<String>();
    private static ArrayList<Integer> ageList = new ArrayList<Integer>();
    
    private static final int cYear = 117; //It is imperative to change these values to what today reflects
    private static final int cMonth = 8;
    private static final int cDay = 5;
    
    private static int gYear;
    private static int gMonth;
    private static int gDay;
    
    public static void setExcelCellstoArr() {
    	try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			@SuppressWarnings("resource")
			Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next(); //Iterates through each cell
                    //getCellTypeEnum shown as deprecated for version 3.15
                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                    if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                    	DataFormatter df = new DataFormatter(); //Lines 38 & 39 convert the numerical value of the cell into an actual string. 
                    	String value = df.formatCellValue(currentCell);
                    	bDList.add(value); //Instead of returning an array list, it adds the values into an array list.
                        //System.out.print(value); //prints out all of the dates. This was used for testing purposes.
                    }

                }
                //System.out.println();

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    public static Integer formulateAge(int gYear, int gMonth, int gDay) { //The formula used to convert a birth date into age in years.
    	return ((((cYear-gYear)*365)+((cMonth-gMonth)*30)+ (cDay-gDay))/365);
    }
    
    public static void setAgeList(){ 
    	ageList = new ArrayList<Integer>();
    	int excelCellsSize = bDList.size(); //Calculates the size of the array list before the for loop is run. In theory it should improve the efficiency of the algorithm.
    	//test = new ArrayList<String>();
    	//test.add("8/6/99");
    	//test.add("12/14/98");
    	for (int index = 0; index < excelCellsSize; index++){ //Iterates through every Date of Birth.
    		String cIndex = bDList.get(index); //Stores "4/11/21" as a string.
    		gYear = Integer.parseInt(cIndex.substring(cIndex.length()-2,cIndex.length())); //Stores the last two string values as integer values.
    		/**
    		 Two if statements, with two nested if statements each. Since there are only four possible ways that the date could be 
    		 formatted (00/00/00 or 00/0/00 or 0/00/00 or 0/0/00) I check to see which format the date falls under, and then I extract
    		 the month and day values from them. 
    		 */
    		if (cIndex.indexOf("/")== 1) {
    			gMonth =  Integer.parseInt(cIndex.substring(0, 1));
    			if (cIndex.lastIndexOf("/")== 4) {
    				gDay = Integer.parseInt(cIndex.substring(2,4));
    			} else if(cIndex.lastIndexOf("/")==3) {
    				gDay = Integer.parseInt(cIndex.substring(2, 3));
    			}
    		} else if (cIndex.indexOf("/")== 2) {
    			gMonth =  Integer.parseInt(cIndex.substring(0, 2));
    			if (cIndex.lastIndexOf("/")== 5) {
    				gDay = Integer.parseInt(cIndex.substring(3,5));
    			} else if(cIndex.lastIndexOf("/")==4) {
    				gDay = Integer.parseInt(cIndex.substring(3,4));
    			}
    		}
    		ageList.add(formulateAge(gYear, gMonth, gDay));
    	}
    }//setAgeList()
    
    public static void main(String[] args) {
    	setExcelCellstoArr();
    	setAgeList();
    	System.out.println("The birthdates: ");
    	System.out.println(bDList);
    	System.out.println("The birth dates converted into years: ");
    	System.out.println(ageList); 
    	int firstR = 0;
    	int secondR = 0;
    	int thirdR = 0;
    	int fourthR = 0;
    	int fifthR = 0;
    	int sixthR = 0;
    	for (int index = 0; index < ageList.size();index++) { //Used to find how many voters fit within certain age groups.
    		int currentAge = ageList.get(index);
    		if (currentAge >= 18 && currentAge <=22)
    			firstR++;
    		else if (currentAge >= 23 && currentAge <= 26)
    			secondR++;
    		else if (currentAge >= 27 && currentAge <= 30)
    			thirdR++;
    		else if (currentAge >= 30 && currentAge <= 35)
    			fourthR++;
    		else if (currentAge >= 31 && currentAge <= 36)
    			fifthR++;
    		else 
    			sixthR++;
    	}//for loop.
    	System.out.println("Age group of Voters||Number of Voters");
    	System.out.println("=====================================");
    	System.out.println("       18-22:      ||    " + firstR);
    	System.out.println("=====================================");
    	System.out.println("       23-26:      ||    " + secondR);
    	System.out.println("=====================================");
    	System.out.println("       27-30:      ||    " + thirdR);
    	System.out.println("=====================================");
    	System.out.println("       30-35:      ||    " + fourthR);
    	System.out.println("=====================================");
    	System.out.println("       31-36:      ||    " + fifthR);
    	System.out.println("=====================================");
    	System.out.println("      Others:      ||    " + sixthR);
    	System.out.println("=====================================");
    }//main
}//BtoD
