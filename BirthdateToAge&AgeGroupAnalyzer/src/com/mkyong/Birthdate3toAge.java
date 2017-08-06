package com.mkyong;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

//This is the first version of my algorithm, and it is highly inefficient. Uses 3 for loops.
public class Birthdate3toAge { //3 for loops, so B3toA.

    private static final String FILE_NAME = "C:\\Users\\gabri\\Documents\\GitHub\\Excel Reader\\excelFiles\\VotingInfoDoc1.xlsx";
    private static ArrayList<String> bDList;
    private static ArrayList<Integer> ageList;
    
    private static final int cYear = 117; //It is imperative to change these values to what today reflects
    private static final int cMonth = 8;
    private static final int cDay = 5;
    
    public static ArrayList<String> getExcelCells() {
    	try {
    		bDList = new ArrayList<String>();
    		
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
                    	bDList.add(value);
                        //System.out.print(value); //prints out all of the dates
                    }

                }
                //System.out.println();

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    	//System.out.println(bDList);
    	//System.out.println(bDList.size());
    	return bDList;
    }
    
    public static Integer formulateAge(int gYear, int gMonth, int gDay) {
    	return ((((cYear-gYear)*365)+((cMonth-gMonth)*30)+ (cDay-gDay))/365);
    }
    
    public static ArrayList<Integer> returnAgeList(){
    	ageList = new ArrayList<Integer>();
    	for (int index = 0; index < getExcelCells().size(); index++){ //Iterates through every Date of Birth.
    		String cIndex = getExcelCells().get(index); //Stores "4/11/21" as a string.
    		int gYear = Integer.parseInt(cIndex.substring(cIndex.length()-2,cIndex.length())); //Stores the last two string values as integer values.
    		
    		//Stores the month value into an integer. It adds string values to a gMonthAsString until it finds a "/"
    		boolean slash = false;
    		String gMonthAsString = "";
    		for (int jindex=0; jindex<cIndex.length(); jindex++) {
        		if (slash == false) {
        			if (cIndex.indexOf("/")== jindex) {
        				slash = true;
        			}
        			else
        				gMonthAsString += cIndex.substring(jindex,jindex+1);
        		}//if statement
        	} //for loop of gMonth
    		int gMonth = Integer.parseInt(gMonthAsString);
    		
    		//Stores the day value into an integer. It doesn't add string values until it finds a "/". Then, it deletes the last three values, which will always be "/00"
    		int postDash = 0;
    		String gDayAsString = "";
    		for (int jindex=0; jindex<cIndex.length(); jindex++) {
	    		if (cIndex.indexOf("/")== jindex) {
	    			postDash++;
	    		} else if (postDash==1) {
	    			gDayAsString += cIndex.substring(jindex, jindex+1);
	    		}
    		}
    		gDayAsString=gDayAsString.replace(gDayAsString.substring(gDayAsString.length()-3, gDayAsString.length()), ""); //Here it deletes the last three values, leaving only the day value
    		int gDay = Integer.parseInt(gDayAsString);
    		
    		ageList.add(formulateAge(gYear, gMonth, gDay));
    		System.out.println(index);
    	}//for loop
    	
    	return ageList;
    }//public static ArrayList<Integer> returnAgeList()
    
    public static void main(String[] args) {
    	returnAgeList();
    }
}