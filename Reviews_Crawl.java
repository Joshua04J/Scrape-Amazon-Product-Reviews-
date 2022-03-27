package Scripts;

import static org.testng.Assert.fail;
import java.time.Duration;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Random;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

public class Reviews_Crawl {

	
	
	WebDriver driver; 
	int i = 0, k = 1;
	XSSFWorkbook wbook; 
	String Infile, Outfile,file ;
	XSSFSheet sh;	
	XSSFRow HeaderRow;
	XSSFCell cell;
	boolean NextPage = true;
	
	
	// *********************************************************

	@BeforeSuite
	public void GetExcelBook() throws IOException {
		
		file = "C:\\Users\\91979\\eclipse-workspace\\Web Scrape Rough\\Triall.xlsx";
		wbook = new XSSFWorkbook(file);
		
		int leftLimit = 97; // letter 'a'
	    int rightLimit = 122; // letter 'z'
	    int targetStringLength = 10;
	    Random random = new Random();

	    String generatedString = random.ints(leftLimit, rightLimit + 1)
	      .limit(targetStringLength)
	      .collect(StringBuilder::new, StringBuilder::appendCodePoint, StringBuilder::append)
	      .toString();
	    System.out.println("Sheet name : " +generatedString);
	    sh =  wbook.createSheet(generatedString);
		
		String[] headers = {"S.no", "Profile Name", "Rating", "Review Title","Review Text (Comment)", "Review Date"};
		// Create 1st row
		for (int heads = 0 ; heads < headers.length ; heads++) {
		try {
			XSSFRow row2 = sh.getRow(0);
			//XSSFCell cell = row2.getCell(heads);
			cell.setCellValue(headers[heads]);
		} catch (NullPointerException NPE)
		{	
			if (heads==0) {
				Row row2 = sh.createRow(0);
			}
			XSSFRow row2 = sh.getRow(0);
			XSSFCell cell = row2.createCell(heads);
			XSSFCell cell2 = row2.createCell(heads+1);
			cell.setCellValue(headers[heads]); // headers[heads]);
			//cell.setCellValue("Null pointer error");
		} 
		}	
	}


	@Test (priority =  1)
	public void OpenBrowser() throws InterruptedException {
		
		System.setProperty("webdriver.chrome.driver", "F:\\sele\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		String Link = "";

		driver.get("https://www.amazon.in/OnePlus-Nord-Bahamas-128GB-Storage/product-reviews/B09RG5R5FG/ref=cm_cr_dp_d_show_all_btm?ie=UTF8&reviewerType=all_reviews");
		//driver.get(Link);
		
		JavascriptExecutor js = (JavascriptExecutor) driver;
		
		try {
		WebElement element=driver.findElement(By.xpath("//*[@class ='a-section a-spacing-large celwidget']"));
	    js.executeScript("arguments[0].remove();", element);
		} catch (Exception NoSuchElement) {
			System.out.println("Exception while eliminating unneccessary Review");
		}
	}


	@Test (priority = 2) 
	public void Get_Reviews() {

		int whilec = 0;
		

		
		try {
			while(NextPage) {
				
				// Actual Review count
				List<WebElement> review =  driver.findElements(By.xpath("//*[@class='a-section review aok-relative']"));
				// Profile name
				List<WebElement> Prof_Names = driver.findElements(By.xpath("//*[@class='a-profile-content']"));
				// Ratings 
				//List<WebElement> Rating = driver.findElements(By.xpath("//*[@class='a-icon a-icon-star a-star-4 review-rating']"));
				// Review Title
				List<WebElement> Titles =  driver.findElements(By.xpath("//*[@class='a-size-base a-link-normal review-title a-color-base review-title-content a-text-bold']"));
				// Review Text
				List<WebElement> Texts = driver.findElements(By.xpath("//*[@class='a-size-base review-text review-text-content']"));
				// Review Date
				List<WebElement> Dates = driver.findElements(By.xpath("//*[@class='a-size-base a-color-secondary review-date']"));


				for (WebElement item  : Texts) {
					for (int j = 0 ; j <= 5 ; j++) {

						// *********************************************************				

						try {
							XSSFRow row2 = sh.getRow(k);
							XSSFCell cell = row2.getCell(j);
							cell.setCellValue("Try");

						} catch (NullPointerException NPE)
						{	
							if (j==0) {	
								Row row2 = sh.createRow(k);
							}
							XSSFRow row2 = sh.getRow(k);
							Cell cell = row2.createCell(j);
							Cell cell2 = row2.createCell(j+1);
							cell.setCellValue("Null pointer error");
						} 

						// *********************************************************


						switch (j) { 
						case 0:		// S.no
							sh.getRow(k).getCell(j).setCellValue(k);
							break;
						case 1:		 	// Profile Names	
							sh.getRow(k).getCell(j).setCellValue(Prof_Names.get(i).getText());
							break;
						case 2:		// Ratings
							sh.getRow(k).getCell(j).setCellValue("Not Available");
							break;
						case 3:		// Review Title
							sh.getRow(k).getCell(j).setCellValue(Titles.get(i).getText());
							break;
						case 4:		// Review Text
							sh.getRow(k).getCell(j).setCellValue(Texts.get(i).getText());
							break;
						case 5:		// Review Date
							sh.getRow(k).getCell(j).setCellValue(Dates.get(i).getText());
						}			// end of case
					}			// end of J loop 
					i = i+1;
					k = k+1;
				}   // each loop ends

				
				whilec = review.size();
				
				try {
					driver.findElement(By.xpath("//*[@class='a-disabled a-last']"));
					NextPage = false;
					
				}catch (Exception Yes_Next) {
					NextPage = true;
				}
				//driver.findElement(By.xpath("//*[@class='a-disabled a-last']"))
				driver.findElement(By.xpath("//*[@class='a-last']")).click();
				Thread.sleep(1500);
				i = 0;

				review.clear();
				Prof_Names.clear();
				Titles.clear();
				Texts.clear();
				Dates.clear();	

			} // while loop ends
		}catch (Exception E) {
			System.out.println("Exception while finding actual Review");
		}
	}

	

	@AfterTest
	public void afterSuite() throws IOException {

		File pate = new File(file);
		String path = pate.getParent();

		
		
		FileOutputStream outFile = new FileOutputStream(new 
				File(path + "\\Product Reviews.xlsx"));
		wbook.write(outFile);
		outFile.close();
		wbook.close(); 
		driver.quit(); 



	}

}
