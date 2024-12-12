package com.Task.Beats;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class App {
  public static void main(String[] args) throws InterruptedException{
    //System.out.println("Hello World!");
    readFromExcel();
  }
  
  private static void readFromExcel() {
	// TODO Auto-generated method stub
	  try {
		  // file input stream file creation
		  FileInputStream file = new FileInputStream(new File ("4BeatsAssignment.xlsx"));
		  
		  // create workbook instance holding refer to .xlsx
		  XSSFWorkbook workbook = new XSSFWorkbook(file);
		  
		  // in assesment xl file have datesheet tabs 
		  // to get those name 
		  String dayName = getDayName();
		  
		  // now get desire sheet from the workbook
		  XSSFSheet sheet = workbook.getSheet(dayName);
		  
		//Iterate through each rows one by one
          Iterator<Row> rowIterator = sheet.iterator();
          while (rowIterator.hasNext()) {
              Row row = rowIterator.next();
              
              // here my condition for cell collection
              String cellValue = row.getCell(2)!=null?row.getCell(2).getStringCellValue():null;
              if(cellValue!=null){
            	  
                  List<String> result=googleSearch(cellValue);                
              	
                  row.createCell(3).setCellValue(getMaxResult(result));
                  row.createCell(4).setCellValue(getMinResult(result));
              }
		  
		  
	  }
          file.close();
          FileOutputStream out = new FileOutputStream(new File ("4BeatsAssignment.xlsx"));
          
          workbook.write(out);
          out.close();
          workbook.close();
	  }
        catch (Exception e) {
            // Display the exception along with line number
            // using printStackTrace() method
            e.printStackTrace();
	}
}

private static String getDayName() {
	
	// TODO Auto-generated method stub
	
	Date date = new Date();
	SimpleDateFormat sdf = new SimpleDateFormat("EEEE", Locale.ENGLISH);
	sdf.setTimeZone(TimeZone.getTimeZone("Asia/Kolkata"));
	return sdf.format(date);
}

private static String getMinResult(List<String> serachResult) {
	// TODO Auto-generated method stub
	// Find the minimum string using a lambda expression
    String minString = serachResult.stream().min((str1, str2) ->
            Integer.compare(str1.length(), str2.length())).orElse(null);
    System.out.println("Minimum String : "+ minString);	      
    return minString;
}

private static String getMaxResult(List<String> serachResult) {
	// TODO Auto-generated method stub
	// Find the maximum string using a lambda expression
    String maxString = serachResult.stream().max((s1, s2) ->
            Integer.compare(( s1).length(), (s2).length())).orElse(null);
    System.out.println("Maximum String : " + maxString);
    return maxString;
}



private static List<String> googleSearch(String string) throws InterruptedException{
		// create a web driver where to search 
		/*
		 * ChromeOptions options = new ChromeOptions();
		 * options.addArguments("--remote-allow-origins=*");
		 */
	
		WebDriver driver = new ChromeDriver();
		
		//driver need what to search
		driver.get("https://www.google.com");
		
		// a browser need to be manage how he react and show 
		driver.manage().window().maximize();
		
		//now he need to do his work where to search
		// and find the searchbox keyword to search the element
		driver.findElement(By.xpath("//*[@title='সার্চ করুন']")).sendKeys(string);
		Thread.sleep(1000);
		
		// Web element search 
		List<WebElement> searchOptions = driver.findElements(By.xpath("//ul[@role='listbox']/li"));
		
		// need list of result 
		List<String> resultList = new ArrayList<String>();
		
		// use forloop to get those options
		for(WebElement option : searchOptions) {
			resultList.add(option.getText());
		}
		// close the driver
		driver.close();
		return resultList;
		
		
		
		
	}
}
