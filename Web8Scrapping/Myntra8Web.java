package Web8Scrapping;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import CommonUtility.BlinkitId;

public class Myntra8Web {

	 public static void main(String[] args) throws Exception{
	        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
	        WebDriver driver = new ChromeDriver();

	        int count = 0;
	        // int finalSp;
	          String spValue = "";
	          String finalSp = "";
	          String offerValue = "NA";
	          String newBName = null;
	          String newPName = null;
	          String newName= null;
	          String mrpValue = null;
	          String originalMrp1 = " ";
	          String originalMrp2 = " ";
	          String originalMrp3 = " ";
	          String originalSp1 = " ";
	          String originalSp2 = " ";
	          String NewAvailability1 = " ";
	        try {
	            // Read URLs from Excel file
	            String filePath = ".\\input-data\\Website 8 Input data.xlsx";
	            FileInputStream file = new FileInputStream(filePath);
	            Workbook urlsWorkbook = new XSSFWorkbook(file);
	            Sheet urlsSheet = urlsWorkbook.getSheet("Myntra");
	            int rowCount = urlsSheet.getPhysicalNumberOfRows();

		            List<String> inputPid = new ArrayList<>(),InputCity = new ArrayList<>(),InputName = new ArrayList<>(),InputSize = new ArrayList<>(),NewProductCode = new ArrayList<>(),
		            		uRL = new ArrayList<>(),UOM = new ArrayList<>(),Mulitiplier = new ArrayList<>(),Availability = new ArrayList<>(),Pincode = new ArrayList<>(),NameForCheck = new ArrayList<>();
		            
	            // Extract URLs from Excel
	            for (int i = 0; i < rowCount; i++) {
	                Row row = urlsSheet.getRow(i);
	                
	                
	                  if (i == 0) {
	                continue;
	            }    
	                
	                
	                Cell inputPidCell = row.getCell(0);
	                Cell inputCityCell = row.getCell(1);
	                Cell inputNameCell = row.getCell(2);
	                Cell inputSizeCell = row.getCell(3);
	                Cell newProductCodeCell = row.getCell(4);
	                Cell urlCell = row.getCell(5);
	                Cell uomCell = row.getCell(6);
	                Cell multiplierCell = row.getCell(7);
	                Cell availabilityCell = row.getCell(8);
	                Cell pinCodeCell = row.getCell(9);        
	                Cell oldNameCell = row.getCell(10);    
	                
	            //   Cell urlCell = row.getCell(0);
	              //  Cell urlCell = row.getCell(0);
	               // Cell idCell = row.getCell(1);
	               // Cell offerCell = row.getCell(2);
	                
	                if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
	                    String url = urlCell.getStringCellValue();
	                    String id = (inputPidCell != null && inputPidCell.getCellType() == CellType.STRING) ? inputPidCell.getStringCellValue() : "";
	                    String city = (inputCityCell != null && inputCityCell.getCellType() == CellType.STRING) ? inputCityCell.getStringCellValue() : "";
	                    String name = (inputNameCell != null && inputNameCell.getCellType() == CellType.STRING) ? inputNameCell.getStringCellValue() : "";
	                    String size = (inputSizeCell != null && inputSizeCell.getCellType() == CellType.STRING) ? inputSizeCell.getStringCellValue() : "";
	                    String productCode = (newProductCodeCell != null && newProductCodeCell.getCellType() == CellType.STRING) ? newProductCodeCell.getStringCellValue() : "";
	                    String uom = (uomCell != null && uomCell.getCellType() == CellType.STRING) ? uomCell.getStringCellValue() : "";
	                    String mulitiplier = (multiplierCell != null && multiplierCell.getCellType() == CellType.STRING) ? multiplierCell.getStringCellValue() : "";
	                    String availability = (availabilityCell != null && availabilityCell.getCellType() == CellType.STRING) ? availabilityCell.getStringCellValue() : "";
	                    String locationSet = (pinCodeCell != null && pinCodeCell.getCellType() == CellType.STRING) ? pinCodeCell.getStringCellValue() : "";
	                    String namecheck = (oldNameCell != null && oldNameCell.getCellType() == CellType.STRING) ? oldNameCell.getStringCellValue() : "";
	                    
	                    inputPid.add(id);
	                    InputCity.add(city);
	                    InputName.add(name);
	                    InputSize.add(size);
	                    NewProductCode.add(productCode);
	                    uRL.add(url);
	                    UOM.add(uom);
	                    Mulitiplier.add(mulitiplier);
	                    Availability.add(availability);
	                    Pincode.add(locationSet);
	                    NameForCheck.add(namecheck);
	                    
						/*
						 * uRL.add(url); ids.add(id); offers.add(offer);
						 */
	                    
	                }
	            }
	            // Create Excel workbook for storing results
	            Workbook resultsWorkbook = new XSSFWorkbook();
	            Sheet resultsSheet = resultsWorkbook.createSheet("Results");

	            Row headerRow = resultsSheet.createRow(0);
	            
	            
	            headerRow.createCell(0).setCellValue("InputPid");
	            headerRow.createCell(1).setCellValue("InputCity");
	            headerRow.createCell(2).setCellValue("InputName");
	            headerRow.createCell(3).setCellValue("InputSize");
	            headerRow.createCell(4).setCellValue("NewProductCode");
	            headerRow.createCell(5).setCellValue("URL");
	            headerRow.createCell(6).setCellValue("Name");
	            headerRow.createCell(7).setCellValue("MRP");
	            headerRow.createCell(8).setCellValue("SP");
	            headerRow.createCell(9).setCellValue("UOM");
	            headerRow.createCell(10).setCellValue("Multiplier");
	            headerRow.createCell(11).setCellValue("Availability");
	            headerRow.createCell(12).setCellValue("Offer");
	            headerRow.createCell(13).setCellValue("Commands");
	            headerRow.createCell(14).setCellValue("Remarks");
	            headerRow.createCell(15).setCellValue("Correctness");
	            headerRow.createCell(16).setCellValue("Percentage");
	            headerRow.createCell(17).setCellValue("Name");
	            headerRow.createCell(18).setCellValue("Name Check");
	            
	            int rowIndex = 1;
	            int headercount = 0;
	            String currentPin = null;
	            
	            for (int i = 0; i < uRL.size(); i++) {
	                String id = inputPid.get(i);
	                String city = InputCity.get(i);
	                String name = InputName.get(i);
	                String size = InputSize.get(i);
	                String productCode = NewProductCode.get(i);
	                String url = uRL.get(i);
	                String uom = UOM.get(i);
	                String mulitiplier = Mulitiplier.get(i);
	                String availability = Availability.get(i);
	                String locationSet = Pincode.get(i);
	                String namecheck = NameForCheck.get(i);
	                
	                
	                
	                try {
	                	
	                	  if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
	                          // Set "NA" values in all three columns
	                          Row resultRow = resultsSheet.createRow(rowIndex++);
	                          resultRow.createCell(0).setCellValue(id);
	                          resultRow.createCell(1).setCellValue(city);
	                          resultRow.createCell(2).setCellValue(name);
	                          resultRow.createCell(3).setCellValue(size);
	                          resultRow.createCell(4).setCellValue(productCode);
	                          resultRow.createCell(5).setCellValue(url);
	                          resultRow.createCell(6).setCellValue("NA");
	                          resultRow.createCell(7).setCellValue("NA");
	                          resultRow.createCell(8).setCellValue("NA");
	                          resultRow.createCell(9).setCellValue("NA");
	                          resultRow.createCell(10).setCellValue("NA");
	                          resultRow.createCell(11).setCellValue("NA");
	                          
	                          System.out.println("Skipped processing for URL: " + url);
	                          continue; // Skip to the next iteration
	                      }
	                	  
	                	  System.out.println("=============="+locationSet+"==============");
	                	  
	                    driver.get(url);
	                    driver.manage().window().maximize();
	                    Thread.sleep(5000);
	                    try {
	                    	
	                        WebElement brandElement = driver.findElement(By.xpath("//h1[@class='pdp-title']"));
	                        WebElement pRODElement = driver.findElement(By.xpath("//h1[@class='pdp-name']"));
	                        newBName = brandElement.getText();
	                        newPName = pRODElement.getText();
	                        newName=newBName+ ""+newPName;
	                        System.out.println(newBName+ ""+newPName);
	                        
	                        }
	                        
	                        catch(NoSuchElementException e) {
	                        	
	                        	WebElement brandElement = driver.findElement(By.xpath("/html/body/div[1]/div/div[3]/div[1]/div[2]/div[2]/div/div[1]/h1/span"));
	                        	 WebElement pRODElement = driver.findElement(By.xpath("/html/body/div[2]/div/div[1]/main/div[2]/div[2]/div[1]/h1[2]"));
	                        	newBName = brandElement.getText();
		                        newPName = pRODElement.getText();
		                        newName=newBName+ ""+newPName;
		                        System.out.println(newBName+ ""+newPName);
	                        	
	                        }
	                        System.out.println("headercount = " + headercount);
	                        
	                        headercount++;
	                        Thread.sleep(2000);
	                        try {
	                            WebElement mrp = driver.findElement(By.xpath("//span[@class='pdp-mrp']//s"));
	                            originalMrp1 = mrp.getText();
	                            mrpValue = originalMrp1.replace("₹", "");
	                            System.out.println(mrpValue);
	                            
	                            } 
	                            
	                            catch(NoSuchElementException e){ 
	                            		try {
	                            		WebElement mrp = driver.findElement(By.xpath("/html/body/div[2]/div/div[1]/main/div[2]/div[2]/div[1]/div/p[1]/span[2]/s/text()[2]"));
	                                    originalMrp2 = mrp.getText();
	                                    mrpValue = originalMrp2.replace("₹", "");
	                                   // mrpValue = originalMrp2;                      
	                                    System.out.println(mrpValue);
	                            		}
	                            		catch (Exception q) {        //span[@class='pdp-price']//strong
	                            			WebElement mrp = driver.findElement(By.xpath("//span[@class='pdp-price']//strong"));
	        	                            originalMrp1 = mrp.getText();
	        	                            mrpValue = originalMrp1.replace("₹", "");
	        	                            System.out.println(mrpValue);
										}
	                                
	                            }
	                        Thread.sleep(2000);	
	                        try {
	                            WebElement sp = driver.findElement(By.xpath(" //span[@class='pdp-price']//strong"));
	                            originalSp1 = sp.getText();
	                            spValue =  originalSp1.replace("₹", "");
	                            System.out.println(spValue);
	                        }
	                        catch(NoSuchElementException s) {
	                        	 WebElement sp = driver.findElement(By.xpath("/html/body/div[2]/div/div[1]/main/div[2]/div[2]/div[1]/div/p[1]/span[1]/strong"));
		                            originalSp1 = sp.getText();
		                            spValue =  originalSp1.replace("₹", "");
		                            System.out.println(spValue);
	                        }
	                        
	                        if(mrpValue.equals(spValue)){
	 	                	   offerValue = "NA";
	 	                   }
	 	                   else {
	 	                   try {
	 	                	   WebElement offer = driver.findElement(By.xpath("//span[@class='pdp-discount']"));
	 	                       String originalOffer = offer.getText();
	 	                       offerValue = originalOffer.replace("-","").replace("% OFF","% Off");
	 	                       
	 	                          System.out.println(offerValue);
	 	                      
	 	                      }catch (Exception e) {
	 	                    	  try {
	 	                    	 WebElement offer = driver.findElement(By.xpath("/html/body/div[2]/div/div[1]/main/div[2]/div[2]/div[1]/div/p[1]/span[3]"));
		 	                       String originalOffer = offer.getText();
		 	                       offerValue = originalOffer.replace("-","").replace("% OFF","% Off");
	 	                    	  }
	 	                    	  catch (Exception m) {
	 	                    		 offerValue ="NA";
								}
		 	                       
							}
	 	                   
	                          
	                       
	                }
	                        //Out Of Stocks
	                        
	                        
	                       
	     					
	                     	 int  result = 1;
	     					try {
	     						
	     					String xpathForMyntra = "//div[@class='pdp-add-to-bag pdp-button pdp-flex pdp-center pdp-out-of-stock ']";

	     					
	     					boolean isElementPresent = !driver.findElements(By.xpath(xpathForMyntra)).isEmpty();

	     			        result = isElementPresent ? 0 : 1;
	     			        
	     			        System.out.println(result);
	     					}
	     					catch(Exception e) {
	     						System.out.println(e.getMessage());
	     					}
	     					
	     					
	                        
	     					//int stock = result;
	     					NewAvailability1 = String.valueOf(result);

	                    	//Screenshots 
	                        BlinkitId screenshot = new BlinkitId();
	  	                   try {
	  	       				screenshot.screenshot(driver, "Myntra", id);
	  	       			} catch (Exception e) {
	  	       				e.fillInStackTrace();
	  	       			
	  	       			}
	                     
	                      Row resultRow = resultsSheet.createRow(rowIndex++);
	                      
	                      resultRow.createCell(0).setCellValue(id);
	                      resultRow.createCell(1).setCellValue(city);
	                      resultRow.createCell(2).setCellValue(name); 
	                      resultRow.createCell(3).setCellValue(size);
	                      resultRow.createCell(4).setCellValue(productCode);
	                      resultRow.createCell(5).setCellValue(url);
	                      resultRow.createCell(6).setCellValue(newName);
	                      resultRow.createCell(7).setCellValue(mrpValue);
	                      resultRow.createCell(8).setCellValue(spValue);
	                      resultRow.createCell(9).setCellValue(uom);
	                      resultRow.createCell(10).setCellValue(mulitiplier);
	                      resultRow.createCell(11).setCellValue(NewAvailability1);
	                      resultRow.createCell(12).setCellValue(offerValue);
	                      resultRow.createCell(13).setCellValue(" ");
	                      resultRow.createCell(14).setCellValue(" ");
	                      resultRow.createCell(15).setCellValue(" ");
	                      resultRow.createCell(16).setCellValue(" ");
	                      resultRow.createCell(17).setCellValue(" ");
	                      resultRow.createCell(18).setCellValue(namecheck);
	                     
	                      
	                      System.out.println("Data extracted for URL: " + url);
	                  } catch (Exception e) {
	                      e.printStackTrace();
	                      
	                      Row resultRow = resultsSheet.createRow(rowIndex++);
	                      resultRow.createCell(0).setCellValue(id);
	                      resultRow.createCell(1).setCellValue(city);
	                      resultRow.createCell(2).setCellValue(name);
	                      resultRow.createCell(3).setCellValue(size);
	                      resultRow.createCell(4).setCellValue(productCode);
	                      resultRow.createCell(5).setCellValue(url);
	                      resultRow.createCell(6).setCellValue("NA");
	                      resultRow.createCell(7).setCellValue("NA");
	                      resultRow.createCell(8).setCellValue("NA");
	                      resultRow.createCell(9).setCellValue(uom);
	                      resultRow.createCell(10).setCellValue(mulitiplier);
	                      resultRow.createCell(11).setCellValue(NewAvailability1);
	                      resultRow.createCell(12).setCellValue(offerValue);
	                      resultRow.createCell(13).setCellValue(" ");
	                      resultRow.createCell(14).setCellValue(" ");
	                      resultRow.createCell(15).setCellValue(" ");
	                      resultRow.createCell(16).setCellValue(" ");
	                      resultRow.createCell(17).setCellValue(" ");
	                      resultRow.createCell(18).setCellValue(namecheck);

	                      System.out.println("Failed to extract data for URL: " + url);
	                      
	                  }
	              }
	              try {
	              	// for store the multiple we can use the time to store the multiple files
	                  SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
	                  String timestamp = dateFormat.format(new Date());
	                  String outputFilePath = ".\\Output\\Myntra8" + timestamp + ".xlsx";
	                  
	                  // Write results to Excel file
	                  FileOutputStream outFile = new FileOutputStream(outputFilePath);
	                  resultsWorkbook.write(outFile);
	                  outFile.close();
	                  
	                  System.out.println("Output file saved: " + outputFilePath);
	              } catch (Exception e) {
	                  e.printStackTrace();
	              }
	             
	          } catch (Exception e) {
	              e.printStackTrace();
	          } finally {
	              if (driver != null) {
	              	System.out.println("DoNe DoNe Scraping DoNe");
	                  driver.quit();
	              }
	          }
	      }
	  

}
