package ExcelSFDC;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import jxl.Cell;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.shooting.ShootingStrategies;

public class excelSFDCConnector {
	
	//TODO see if ASHOT can handle the timer for the screenshot
	//TODO update Timer and spaces on front and back of input.
	
	
	public static void main(String[] args) {
		
		//Set up the variables
		WebElement thisOne = null;
		WebElement theNextPage = null;
		Boolean firstTrip = true;
		Boolean foundTheNextPage = false;
		Date today = new Date();
		Boolean foundIt = true;
		SimpleDateFormat dateFormat = new SimpleDateFormat("MM-dd-yy-hh-mm");
		String inputPath = "Resources\\Input\\Instructions.xls";
		String outputPath = "Resources\\Output\\Results"+ dateFormat.format(today)+".xls";
		String baseURL = null;
		
		File inputFile = new File(inputPath);
		File outputFile = new File(outputPath);
				
		//Set up Chrome to be the web browser
		System.setProperty("webdriver.chrome.driver", "Resources\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		
		//open the workbook
		try {
			Workbook inputWorkbook = Workbook.getWorkbook(inputFile);
			WritableWorkbook outputWorkbook = Workbook.createWorkbook(outputFile,inputWorkbook);					
			
			WritableCellFormat format = new WritableCellFormat();
	        format.setBackground(Colour.RED);			
			for (int i = 0; i < outputWorkbook.getNumberOfSheets();i++){
				
				WritableSheet outputSheet = outputWorkbook.getSheet(i);
				
				for (int j = 1; j < outputSheet.getRows(); j++) {

					// get the information from row 1, columns 1 and 2
					// (column,row)
					Cell firstCell = outputSheet.getCell(0, j);
					Cell secondCell = outputSheet.getCell(1, j);

					//added trim to strings to take out whitespace 9/27/2018
					String doWhat = firstCell.getContents().trim();
					String withWhom = secondCell.getContents().trim();
					
					foundIt = false; 
					/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
					//Compares will proved a true or false with two saved variables
					if (doWhat.equals("Compare")) {
						//checking to see if Saved is in withWhom so that we can separate and get the cell to pull the ID from
						if(withWhom.contains("Cells")){

							//This will get the cell range in the form of character/row (A6)
							String[] word = withWhom.split(" ", -2); 
							String cell1 = word[1];
							String cell2 = word[2];
							
							//Going to the sheet we are working on and getting the cell referenced
							Cell getFromCell1 = outputSheet.getCell(cell1);
							Cell getFromCell2 = outputSheet.getCell(cell2);
							
							//Getting the contents of that cell
							String output1 = getFromCell1.getContents();
							String output2 = getFromCell2.getContents();
							
							String true1 = "Same";
							String false1 = "Different";
							
							//adding the id to the end of the url first opened if it is less than or equal 18 characters otherwise it is the whole url							
							if(output1.equals(output2)){
								//print true in Column C
								Label temp = new Label(2, j, true1);
								outputSheet.addCell(temp);
								System.out.println("Both cells are the same");
							
							}else{
								//print false in column c							
								Label temp = new Label(2, j, false1);
								outputSheet.addCell(temp);
								System.out.println("Both cells are Different");
							//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
							} 
						}
					}
					
					if (doWhat.equals("Open")) {
						//checking to see if Saved is in withWhom so that we can separate and get the cell to pull the ID from
						if(withWhom.contains("Saved")){
							
							//This will get the cell range in the form of character/row (A6)
							String getFrom = withWhom.substring(6, withWhom.length());							
							//Going to the sheet we are working on and getting the cell referenced
							Cell getFromCell = outputSheet.getCell(getFrom);
							
							//Getting the contents of that cell
							String contents = getFromCell.getContents();
							
							//adding the id to the end of the url first opened
							driver.get(baseURL + contents);
							
							//adding the id to the end of the url first opened if it is less than or equal 18 characters otherwise it is the whole url							
							if(contents.length() <= 18){
								driver.get(baseURL + contents);
							}else{
								driver.get(contents);
							} 
						}else{
							
							driver.get(withWhom);
							driver.manage().window().maximize();
	
							if (firstTrip == true) {
								try {
									Thread.sleep(60000);
									//setting the base url as the first url that was opened
									baseURL = withWhom;
								} catch (InterruptedException e1) {
	
									e1.printStackTrace();
								}
								firstTrip = false;
							}
						}
					}  else if (doWhat.equals("Click")) {
												
						try{
							do {
								foundIt = false;
								
								if (driver.findElements(By.xpath("//a[text()='Next Page>']")).size() > 0){
									System.out.println("Found Next Page");
																			
									theNextPage = driver.findElement(By.xpath("//a[text()='Next Page>']"));									
					
									foundTheNextPage = true;
									
								}else if(driver.findElements(By.xpath("//a[text()='Next']")).size() > 0){
									System.out.println("Found Next");
									
									theNextPage = driver.findElement(By.xpath("//a[text()='Next']"));
									
									foundTheNextPage = true;
								}
								
								
						
								if (driver.findElements(By.xpath("//input[@title='" + withWhom + "']")).size() > 0) {
									System.out.println(withWhom + " input By Title");
									thisOne = driver.findElement(By.xpath("//input[@title='" + withWhom + "']"));
									thisOne.click();
									foundIt = true;
	
								} else if (driver.findElements(By.xpath("//input[@value='" + withWhom + "']")).size() > 0) {
									
									System.out.println(withWhom + " input by Value");
									
									//Changing to allow iteration through to find the displayed one 9/27/2018
									List <WebElement> lookingFor = driver.findElements(By.xpath("//input[@value = 'OK']"));
									
									for (WebElement each: lookingFor){
													
										if(each.isDisplayed()){
											each.click();
										}
									}
									
									foundIt = true;
									
								} else if (driver.findElements(By.xpath("//a[@title='" + withWhom + "']")).size() > 0) {
									System.out.println(withWhom + " anchor by title");
									thisOne = driver.findElement(By.xpath("//a[@title='" + withWhom + "']"));
									thisOne.click();
									foundIt = true;

								} else if (driver.findElements(By.xpath("//a[text()='" + withWhom + "']")).size() > 0) {
									System.out.println(withWhom + " anchor by text");
	
									thisOne = driver.findElement(By.xpath("//a[text()='" + withWhom + "']"));
									Actions build = new Actions(driver); 
									build.moveToElement(thisOne).build().perform();
									WebElement tryAgain = driver.findElement(By.xpath("//a[text()='" + withWhom + "']"));
									tryAgain.click();
									foundIt = true;

								} else if (driver.findElements(By.xpath("//a[@value='" + withWhom + "']")).size() > 0) {
									System.out.println(withWhom + " anchor by value");
									thisOne = driver.findElement(By.xpath("//a[value()='" + withWhom + "']"));
									thisOne.click();
									foundIt = true;
									
								} else if (driver.findElements(By.xpath("//img[@alt='" + withWhom + "']")).size() > 0) {
									System.out.println(withWhom + " image by alt");
									thisOne = driver.findElement(By.xpath("//img[@alt='" + withWhom + "']"));
									thisOne.click();
									foundIt = true;
	
								} else if (driver.findElements(By.xpath("//span[text()='" + withWhom + "']")).size() > 0) {
									System.out.println(withWhom + " span by text");
	
									//adding a wait to make sure that Contacts get a chance to fill in 
									driver.manage().timeouts().implicitlyWait(1,TimeUnit.SECONDS);
									thisOne = driver.findElement(By.xpath("//span[text()='" + withWhom + "']"));
									
									Actions build = new Actions(driver);
									build.moveToElement(thisOne).build().perform();
									WebElement tryAgain = driver.findElement(By.xpath("//span[text()='" + withWhom + "']"));
									tryAgain.click();
									foundIt = true;
									
								}else{
									theNextPage.click();
							}
						} while(foundIt == false && foundTheNextPage == true);
							
						}catch (Exception e){
							errorConditional(format, outputSheet, j, withWhom);
						}

					} else if (doWhat.equals("Select")) {
						try{
							if (driver.findElements(By.xpath("//label[text()='" + withWhom + "']")).size() > 0) {
								System.out.println(withWhom + " By Label");
								thisOne = driver.findElement(By.id(driver
										.findElement(By.xpath("//label[text()='" + withWhom + "']")).getAttribute("for")));

							} else if (driver.findElements(By.xpath("//input[@title='" + withWhom + "']")).size() > 0) {
								System.out.println(withWhom + " By Input");
								thisOne = driver.findElement(By.xpath("//input[@title='" + withWhom + "']"));

							} else if (driver.findElements(By.xpath("//label[text()='" + withWhom + "']")).size() > 0) {
								System.out.println(withWhom + " By label");
								thisOne = driver.findElement(By.id(driver
										.findElement(By.xpath("//label[text()='" + withWhom + "']")).getAttribute("for")));

							}
							//Created 8/23/2018 for Search to find Id for the Search bar
							else if (driver.findElements(By.xpath("//input[@id='"+withWhom+"']")).size() >0){ 
		                		System.out.println(withWhom + " By input");                		
		                		thisOne = driver.findElement(By.xpath("//input[@id='"+withWhom+"']"));
		                		thisOne.click(); 
	                	
							}
						}catch (Exception e) {
								errorConditional(format, outputSheet, j, withWhom);
							}

						} else if (doWhat.equals("Enter")) {

							try{
								/////////////////////////////////////////////////////////////////////////////////////////////////////////
								if(withWhom.contains("Saved")){								
									//This will get the cell range in the form of character/row (A6)
									String getFrom = withWhom.substring(6, withWhom.length());							
									//Going to the sheet we are working on and getting the cell referenced
									Cell getFromCell = outputSheet.getCell(getFrom);
									
									//Getting the contents of that cell
									String contents = getFromCell.getContents();
									driver.get(contents);
									///////////////////////////////////////////////////////////////////////////////////////////////
								}
								thisOne.clear();
								thisOne.sendKeys(withWhom);
							}catch (Exception e){
								errorConditional(format, outputSheet, j, withWhom);
							}
							
	

						}else if (doWhat.equals("Save")) {
							String url = driver.getCurrentUrl();
							
							if(withWhom.equals("URL")){
								
								Label temp = new Label(2, j, url);
								outputSheet.addCell(temp);
								
							}else if(withWhom.equals("ID")){
								if(url.contains("=")){
									
									int last = url.indexOf("&id=") + 4;//Changed to at &id on 9/11/2018
									String id = url.substring(last,last + 15);									
									
									Label temp = new Label(2, j, id);
									outputSheet.addCell(temp);
									
								}else{
									String id = url.substring(url.length() - 15);			
									Label temp = new Label(2, j, id);
									outputSheet.addCell(temp);
								}
							}
						}else if (doWhat.equals("Pick")) {
							try{
								//Adding a check to see if thisOne is still active otherwise find the one on the visualforce page 9/20/2018
								if(thisOne.isDisplayed()){
									Select dropDown = new Select(thisOne);								
									dropDown.selectByVisibleText(withWhom);
								}
								else{
									//Sleeping 3 seconds to allow for the page to finish loading 9/20/2018
									//Thread.sleep(3000);
									
									//checking to see if the option we are choosing exists 9/20/2018
									if(driver.findElements(By.xpath("//select/option[text() ='" + withWhom +"']")).size() > 0){
										
										//getting that option so we can click and set 9/20/2018
										thisOne = driver.findElement(By.xpath("//select/option[text() ='" + withWhom +"']"));
										thisOne.click();
									}										
																		
								}
								
							}catch (Exception e){
								errorConditional(format, outputSheet, j, withWhom);
							}
							
						}else if (doWhat.equals("Go to")){
							
							if(withWhom.equals("Setup")){
								
								String url = driver.getCurrentUrl();
								int last = org.apache.commons.lang3.StringUtils.ordinalIndexOf(url, "/", 3);
								String base = url.substring(0, last);
								System.out.println("Navigate to Setup in " + base);
								driver.get(base +"/setup/forcecomHomepage.apexp ");						
							}
						}
						// Takes a screenshot and places image in XLS 
						else if (doWhat.equals("Screenshot")){
							
							driver.switchTo().activeElement();
							driver.manage().timeouts().pageLoadTimeout(8, TimeUnit.SECONDS);

							/**try {
								Thread.sleep(10000);
						    }
						    catch (InterruptedException e) {
						    	e.printStackTrace();
								System.out.println("Exception found: " + e);
						    } 	**/
							
							WritableImage sc = screenshot(driver);
							sc.setRow(j);							
							outputSheet.addImage(sc);
							
							// use this to put the name of the pic in column C							
							String screenShotName = (String) sc.getImageFilePath();														
							Label temp = new Label(2, j, screenShotName);
							outputSheet.addCell(temp); 
							
						   
							System.out.println("Screenshot Taken");
						}
						else if (doWhat.equals("Toggle")){
							thisOne.click();
						}
					}
				}
				outputWorkbook.write();
	        	outputWorkbook.close();
	        		
	        	inputWorkbook.close();
	        	driver.quit();
			} catch (BiffException e) {				
				e.printStackTrace();
			} catch (IOException e) {				
				e.printStackTrace();
			} catch (WriteException e) {				
				e.printStackTrace();
			}			
			System.out.println("-----------------------------------------------------"+'\n' + "Automation Completed");
		}

	private static void errorConditional(WritableCellFormat format, WritableSheet outputSheet, int j, String withWhom) {
		System.out.println("Did not find element " + withWhom);

		WritableCell cellOne = outputSheet.getWritableCell(0, j);
		WritableCell cellTwo = outputSheet.getWritableCell(1, j);

		cellOne.setCellFormat(format);
		cellTwo.setCellFormat(format);
	}
	//Screenshot creator
	public static WritableImage screenshot(WebDriver driver) throws IOException, BiffException, RowsExceededException, WriteException {
		
		//variables used to make IMG unique
		Date today = new Date();
		SimpleDateFormat imgTimeStamp = new SimpleDateFormat("MMddYY-hhmmssSS");
				
		//Get the entire page Screenshot
		// adding Ashot jar to take full screenshot must have screen set to 100% any above will miss pieces 9/10/2018
		Screenshot screenshot = new AShot().shootingStrategy(ShootingStrategies.viewportPasting(100)).takeScreenshot(driver);
	    
	    // Copy the screenshot to disk
	    File screenshotLocation = new File("Resources\\IMG\\savedimage_" + imgTimeStamp.format(today) + ".png");		
		
	    //Write Pic to file
	    ImageIO.write(screenshot.getImage(),"PNG",screenshotLocation);
		
	    WritableImage image = new WritableImage(3, 0, 1, 4, screenshotLocation);
	    
	    return image;
	}  

}