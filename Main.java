import java.io.File;
import java.io.IOException;
import java.util.Date; 
import jxl.*;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import jxl.read.biff.BiffException; 
import java.util.Scanner;
import java.util.Date;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

 /*
  * Created by Ellery Kreloff on behalf of Aaron G, to make his SOAR data entry tolerable.
  * 7/14/16
  * JC
  */
  
public class Main {
	
	public static final Scanner scnr = new Scanner(System.in);

	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException, ParseException {
		System.out.print("Enter name for new excel spreadsheet. Include .xls at the end of the name.\n"
				+ "Example: newfile.xls <-- Don't put spaces in it pls. \nName: ");
		String filename = scnr.nextLine();
		WritableWorkbook workbook = Workbook.createWorkbook(new File(filename));
		//get first sheet
		System.out.println("Creating workbook...");
		WritableSheet sheet = workbook.createSheet("Sheet1", 0);
		//Auto Add columns names 
		System.out.println("If your code has the same column names every time, I've provided instructions \n"
				+ "in the code to save you some time in the long run. Just exit the program and \n"
				+ "look at line 27. Otherwise, ignore.");
		
		/* UNCOMMENT THIS
		//LINE 27
		//If you want to have the same column names every time, uncomment the following section of code
		//And fill in the required information.
		String[] columnNames = new String[13];
		//starting at 1 for real reasons I'm not a dumbass I promise
			columnNames[1] = "example name";
			columnNames[2] = "";
			columnNames[3] = "";
			columnNames[4] = "";
			columnNames[5] = "";
			columnNames[6] = "";
			columnNames[7] = "";
			columnNames[8] = "";
			columnNames[9] = "";
			columnNames[10] = "";
			columnNames[11] = "";
			columnNames[12] = "";
		//fill those bad boys out, save
		 
		 //this part adds the column names to the sheet
		 System.out.println("Now adding column names...");
		 for(int i = 1; i <= 12; i++){
		 	Label label = new Label(i, 0, columnNames[i]);
		 	sheet.addCell(label);
		 }
		 UNCOMMENT THIS*/
		
		
		//if you just filled out the previous section, delete or comment the following section starting here:
		System.out.println("Now adding column names...");
		for(int i = 1; i <= 12; i++){
			System.out.print("Enter name for column " + i + ": ");
			String columnName = scnr.nextLine();
			Label label = new Label(i, 0, columnName); 
			sheet.addCell(label); 
		}
		//ending here
		
		//beginning autofill dates 
		System.out.print("Enter an integer of how many days you want this to generate for. \n"
				+ "Ex. if you input 7 it will make 7 rows for 7 days. DON'T enter not an integer \n"
				+ "I was too lazy to do input verification. You will break the program and have to start over.\n"
				+ "You're not a super hacker for breaking my s***, I promise. I'm just a lazy dbag.\n"
				+ "Enter NUMBER: ");
		//where it parses the int passed in
		String daysString = scnr.nextLine();
		int days = Integer.parseInt(daysString);
		
		//I love making computers do all the work for me
		//Iterating through every day + setting the label of the column
		//this section creates all the date rows in column A
		System.out.println("Now creating date columns...");
		for(int k = 1; k <= days; k++){
			//constructing dates fight me
			Date dt = new Date();  // Start date
			//using the calendar because java makes everything difficult
			Calendar c = Calendar.getInstance();
			//what time is it
			c.setTime(dt);
			//this will add one day at a time for each iteration of the for loop
			c.add(Calendar.DATE, k);  // number of days to add
			//the new date
			dt = (c.getTime());  // dt is now the new date
			//make the label
			Label label = new Label(0, k, dt.toString());
			//set the cell
			sheet.addCell(label);
		}
		int rowCounter = 1;
		int columnCounter = 1;
		
		System.out.println("Now commencing data entry...");
		for(int l = 0; rowCounter <= days; l++){
			for(int m = 0; m < 12; m++){
				System.out.print("Enter data for row " + rowCounter + ", column " + columnCounter + ": ");
				String name = scnr.nextLine();
				Label label = new Label(columnCounter, rowCounter, name);
				sheet.addCell(label);
				columnCounter++; 
			}
			columnCounter = 1;
			rowCounter++;
		}
		System.out.println("The file has been created and can now be found in your workspace/excel-spreadsheet-entry\n"
				+ "or whatever you called it directory. It should be next to the bin and src folders. Exiting.");
		workbook.write(); 
		workbook.close();
		
	}

}
