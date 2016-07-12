package org.openqa.selenium.example;


import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import java.io.File;

import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Selenium2Example  {
	
	static WritableWorkbook workbook;
	static WritableSheet sheet;
	static int row = 1, count = 0;
	
    public static void main(String[] args) throws InterruptedException, TimeoutException, IOException, RowsExceededException, WriteException{
        /* Create a new instance of the Firefox driver
           Notice that the remainder of the code relies on the interface, 
           not the implementation. */
        WebDriver driver = new FirefoxDriver();
        callWorkbookInit();

        /*And now use this to visit Google
        driver.get("http://www.google.com");
        Alternatively the same thing can be done like this
        driver.navigate().to("http://www.google.com");*/
        driver.get("http://www.publicnoticeads.com/az/search/advancedsearch.asp");

        /*Find the text input element by its name
        ///WebElement element = driver.findElement(By.name("q"));*/
        WebElement element = driver.findElement(By.name("lstCounties"));
        element.sendKeys("All Counties/Parishes");
        
        element = driver.findElement(By.name("txtSearchWordsAnd"));
        element.sendKeys("maricopa");
        
        element = driver.findElement((By.name("txtSearchWordsOr")));
        element.sendKeys(("deceased, personal, estate"));
        
        element = driver.findElement((By.name("txtDateFrom")));
        element.sendKeys(("01/01/16"));
        
        element = driver.findElement((By.name("txtDateTo")));
        element.sendKeys(("05/21/16"));
        
        /* Now submit the form. WebDriver will find the form for us from the element */
        element.submit();

        /* Check the title of the page */
        System.out.println("Page title is: " + driver.getTitle());
        
        /* Google's search is rendered dynamically with JavaScript.
           Wait for the page to load, timeout after 10 seconds */
        (new WebDriverWait(driver, 60)).until(new ExpectedCondition<Boolean>() {
           public Boolean apply(WebDriver d) {
        	   return d.getPageSource().contains("Public Notice Search Results");
            }
        });
        
        System.out.println("Page loaded successfully");


        // driver.findElement(By.xpath("//table/tbody/tr[3]/td[3]/font/small/a")).click();
        int count = driver.findElements(By.xpath("//table/tbody/tr")).size();
        int itr, numItrs=0;

        while (numItrs < 80)
        {
        	/* Google's search is rendered dynamically with JavaScript.
            Wait for the page to load, timeout after 10 seconds */
	       (new WebDriverWait(driver, 60)).until(new ExpectedCondition<Boolean>() {
	           public Boolean apply(WebDriver d) {
	       	   return d.getPageSource().contains("Public Notice Search Results");
	            }
	        });
         
            for (itr = 2; itr <= count - 2; itr++)
            {
                driver.findElement(By.xpath("//table/tbody/tr[" + itr + "]/td[3]/font/small/a")).click();
                TimeUnit.SECONDS.sleep(1);

                /*  Implement this function to print in Excel */
                printToExcel(driver);

                driver.navigate().back();
                (new WebDriverWait(driver, 60)).until(new ExpectedCondition<Boolean>() {
                    public Boolean apply(WebDriver d) {
                 	   return d.getPageSource().contains("Public Notice Search Results");
                     }
                 });
                
            }
            
            numItrs++;
            System.out.println("Page iteration complete");
            
            if (driver.findElement(By.partialLinkText("Next Records")).isDisplayed())
            {
                driver.findElement(By.partialLinkText("Next Records")).click();   
            }
            else
            {
                break;
                
            }        
            
        }
        workbook.write();
        workbook.close();
        
        /* Close the browser */
        driver.quit();
    }
    
    
    private static void callWorkbookInit() throws IOException, RowsExceededException, WriteException
    {
    	workbook = Workbook.createWorkbook(new File("output.xls"));
    	sheet = workbook.createSheet("First Sheet", 0);
    }
    /**
     * @param driver
     * @throws IOException 
     * @throws WriteException 
     * @throws RowsExceededException 
     */
    private static void printToExcel(WebDriver driver) throws IOException, RowsExceededException, WriteException
    {
    	
    	
//    	Label label = new Label(0, 2, "A label record"); 
//    	sheet.addCell(label); 
//    	
//    	Number number = new Number(3, 4, 3.1459); 
//    	sheet.addCell(number);
    	
//    	WritableWorkbook workbook = Workbook.createWorkbook(new File("output.xls"));
//    	WritableSheet sheet = workbook.createSheet("First Sheet", 0);
    	WebElement pubInfo = driver.findElement(By.id("publicationInfo"));
    	WebElement notText = driver.findElement(By.id("noticeText"));
    	 
    	/* Print to CSV file */
        String info = pubInfo.getText();
        String[] myText = info.split("\\n");
        String info2 = notText.getText().replace(",", " ");
        
        
        String tempCounty = myText[0];
        String tempPrintedIn = myText[1];
        String tempPrintedOn = myText[2];
        
        String county = tempCounty.substring(tempCounty.indexOf(": ")+2);
        String printedIn = tempPrintedIn.substring(tempPrintedIn.indexOf(": ")+2);
        String printedOn = tempPrintedOn.substring(tempPrintedOn.indexOf(": ")+2);
        
        try
        {
       
	    	sheet.addCell(new Label(0, row, county));
	    	
	    	sheet.addCell(new Label(1, row, printedIn));
	    	
	    	sheet.addCell(new Label(2, row, printedOn));
	    	
	    	sheet.addCell(new Label(3, row, info2));
	        
	    	row++;
	    	
//	    	workbook.write();
        }
        catch (WriteException e)
        {
        	System.out.println(e.getMessage());
        }
              		
//        try{
////    		WriteFile data = new WriteFile("demo.csv",true);
////    		data.writeToFile(county + "," + printedIn + "," + printedOn + "," + info2);
//        	Label label = new Label(0, 2, "A label record"); 
//        	sheet.addCell(label);
//    		
//    	}
//    	catch(IOException e){
//    		System.out.println(e.getMessage());
//    	}
        
        /* Print to CSV file */
        
        
    }
}