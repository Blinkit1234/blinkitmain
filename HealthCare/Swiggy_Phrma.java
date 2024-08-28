package HealthCare;

import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;

import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Swiggy_Phrma {

    public static void main(String[] args) throws Exception{
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        int count = 0;
        // int finalSp;
          String spValue = "";
          String finalSp = "";
          String offerValue = "NA";
          String newName = null;
          String brandName=null;
          String pname=null;
          String mrpValue = null;
          String originalMrp1 = " ";
          String originalMrp2 = " ";
          String originalMrp3 = " ";
          String originalSp1 = " ";
          String originalSp2 = " ";
        //  String uomNew = " ";
          String NewAvailability1 = " ";
          
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\Phrma input data.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Swiggy");
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
                
            //    Cell urlCell = row.getCell(0);
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
            headerRow.createCell(12).setCellValue("Commands");
            headerRow.createCell(13).setCellValue("Remarks");
            headerRow.createCell(14).setCellValue("Correctness");
            headerRow.createCell(15).setCellValue("Percentage");
            headerRow.createCell(16).setCellValue("Name");
            headerRow.createCell(17).setCellValue("Offer");
            headerRow.createCell(18).setCellValue("NameForCheck");
            
            int rowIndex = 1;
            int headercount = 0;
            String currentPin =null;
            
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
                          resultRow.createCell(12).setCellValue("NA");
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	  
                    //location sets 
                	  
                	  if (currentPin == null || !currentPin.equals(locationSet)) {
                		  
                		  
                		  driver.get("https://www.swiggy.com/");
                		  driver.manage().window().maximize();
                          
                          //location sets 
                          WebElement location = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/header/div/div/div/span[1]"));
      					location.click();
      					String tempPinNumber = "";
      					for (int j = 0; j < 150; j++) {
      						try {
      							driver.findElement(
      									By.xpath("//*[@id=\"overlay-sidebar-root\"]/div/div/div[2]/div/div/div[2]/div[2]/div/input"))
      									.sendKeys(Keys.ENTER);
      							
      							Thread.sleep(1000);
      							
      							driver.findElement(
      									By.xpath("//*[@id=\"overlay-sidebar-root\"]/div/div/div[2]/div/div/div[2]/div[2]/div/input")).clear();
      							
      							Thread.sleep(1000);
      							
      							System.out.println("print the crt pin number" + locationSet);
      							
      							String crtPin = locationSet;
      							driver.findElement(
      									By.xpath("//*[@id=\"overlay-sidebar-root\"]/div/div/div[2]/div/div/div[2]/div[2]/div/input"))
      									.sendKeys(crtPin);
      							
      							
      							Thread.sleep(1000);
      							
      							currentPin = locationSet;
      							
      						/*	driver.findElement(
      									By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
      									.sendKeys(Keys.ENTER);   */
      							
      							for (int k = 0; k <= 50; k++) {
      								try {
      									tempPinNumber = driver.findElement(By.xpath(
      											"//*[@id=\"overlay-sidebar-root\"]/div/div/div[2]/div/div/div[2]/div[2]/div/input"))
      											.getAttribute("value");
      									if (tempPinNumber.equals(locationSet)) {
      										break;
      									}
      								} catch (Exception e) {
      									if (i == 50) {
      										Assert.fail(e.getMessage());
      									}
      								}
      							}
      							
      							Thread.sleep(1000);
      							
      							driver.findElement(By.xpath("/html/body/div[3]/div/div/div[2]/div/div/div[3]/div/div/div[1]")).click();
      							
      							Thread.sleep(1000);
      							
      							currentPin = locationSet;
      							
      							break;
      						} catch (Exception e) {
      							e.getCause();
      							if (j == 300) {
      								Assert.fail(e.getMessage());
      							}
      						}
      					}
                          }
                    
                	  driver.get(url);
                      driver.manage().window().maximize();
                	  
                	Thread.sleep(500);  
                    try {
                    	WebElement brandElement = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div[2]/div[1]/div[1]"));
                        brandName = brandElement.getText();
                        	
                        WebElement nameElement = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div[2]/div[1]/div[2]"));
                        pname = nameElement.getText();
                        
                        newName=brandName+""+pname;
                         System.out.println(newName);
                    }
                    
                    catch(NoSuchElementException e) {
                    	WebElement brandElement = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div[2]/div[1]/div[1]"));
                        brandName = brandElement.getText();
                        
                    	WebElement nameElement = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div[2]/div[1]/div[2]"));
                    	pname = nameElement.getText();
                    	
                    	newName=brandName+""+pname;
                        System.out.println(newName);
                    	
                    }
                    System.out.println("headercount = " + headercount);
                    
                    headercount++;
                    
                    
                    // Mrp
                    Thread.sleep(500);
                    try {
                    WebElement mrp = driver.findElement(By.xpath("//div[@class='sc-aXZVg fVWuLc _2XPBo _1QyO8']"));
                    originalMrp1 = mrp.getText();
                    mrpValue = originalMrp1.replace("₹", "");
                    System.out.println(mrpValue);
                    
                    }
                    
                    catch(NoSuchElementException e){ 
                    	try {
                    		
                    		WebElement mrp = driver.findElement(By.xpath("//*[@id=\\\"product-details-page-container\\\"]/div/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div"));
                            originalMrp2 = mrp.getText();
                            mrpValue = originalMrp2.replace("₹", "");
                           // mrpValue = originalMrp2;                      
                            System.out.println(mrpValue);
                        
                    }
                    	catch(Exception ex) {
                    		try {
                    		WebElement mrp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]/span[2]/span[2]"));
                           // WebElement mrp = driver.findElement(By.xpath("/html/body/div[2]/div/div[7]/div[3]/div[4]/div[12]/div/div/div[4]/div[2]/span/span[1]/span[2]/span/span[2]"));
                            originalMrp3 = mrp.getText();
                            mrpValue = originalMrp3.replace("₹", "");                      
                            System.out.println(mrpValue);
                    		}
                    		catch(Exception exx) {
                    			mrpValue = "NA";
                    		}
                    		}
                    	}
                    Thread.sleep(500);
                   try {
                    WebElement sp = driver.findElement(By.xpath("//div[@class='sc-aXZVg bzVIAg _2XPBo']"));
                    originalSp1 = sp.getText();
                    spValue =  originalSp1.replace("₹", "");
                    System.out.println(spValue);
                   }
                   catch(Exception e) {
                	   
                	   try {
                	   WebElement sp = driver.findElement(By.xpath("//*[@id=\"__next\"]/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[4]/div[1]/h4"));
                       originalSp2 = sp.getText();
                       spValue =  originalSp2.replace("₹", "");
                       System.out.println(spValue);
                	   }
                       catch(Exception exx) {
                    		   spValue = mrpValue;
            		   
                       }
                   }
                   
                   //Uom scrape
                   
//                   try {
//                       WebElement uomScrap = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]"));
//                       uomNew = uomScrap.getText();
//                       //spValue =  originalSp1.replace("₹", "");
//                       System.out.println(uomNew);
//                      }
//                      catch(Exception e) {
//                    	  
//                    	  uomNew = "NA";
//                   	   //WebElement uomScrap = driver.findElement(By.xpath("//*[@id=\"__next\"]/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[4]/div[1]/h4"));
//                      	//uomNew = uomScrap.getText();
//                          //spValue =  originalSp2.replace("₹", "");
//                          System.out.println(uomNew);
//                      }
//                   
                   // offer
                   try {
                       WebElement offer = driver.findElement(By.xpath("//*[@id=\"product-details-page-container\"]/div/div[2]/div[1]/div[1]/div"));
                       offerValue = offer.getText();
                       
                      /* Pattern pattern = Pattern.compile("\\((.*?)\\)");
                       Matcher matcher = pattern.matcher(offer1);
                       
						if(matcher.find()) { 
							  String offer2 = matcher.group(1);
							  offerValue = offer2.replace("%","% Off");
						  }
                       //offerValue = offer.getText();  */
                       System.out.println(offerValue);
                      }
                      catch(Exception e) {
                    	  try {
                   	   WebElement offer = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]"));
                   	   String NewOffer = offer.getText();
                       offerValue = NewOffer.replace("-","").replace("%","% Off");
                          System.out.println(offerValue);
                    	  }
                    	  catch(Exception ex){
                    		  offerValue = "NA";
                    	  }
                      }    
                   
                   //out Of Stock
                   
                   int result=1;
                   if (url.contains("NA")) {
                	   NewAvailability1 = "NA";
                	   } 
                   else {
                	 
                	   try {
                	   // Define the texts to check for
                		   String[] textsToCheck = {
                				   "Currently Unavailable",
                				   "Currently out of stock in this area.",
                				   "Sold Out"
                				   };

                	   // Get the page source
                	   String pageSource = driver.getPageSource();
                	   boolean isTextPresent = false;

                	   // Check for the presence of any of the texts
                	   for (String text : textsToCheck) {
                	   if (pageSource.contains(text)) {
                	   isTextPresent = true;
                	   break;
                	   }
                	   }

                	   // Determine the result based on the presence of the text
                	   result = isTextPresent ? 0 : 1;
                	   System.out.println(result);
                	   } catch (Exception e) {
                	   System.out.println("Error checking availability: " + e.getMessage());
                	   result = -1;
                	   }
                	   }

                	   // Assign final availability status
                	   NewAvailability1 = String.valueOf(result);
                	   
                   
                   
                   		//Screenshots 
                      BlinkitId screenshot = new BlinkitId();
	                   try {
	       				screenshot.screenshot(driver, "Swiggy", id);
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
                    resultRow.createCell(12).setCellValue(" ");
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(offerValue);
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
                    resultRow.createCell(12).setCellValue(" ");
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(offerValue);
                    resultRow.createCell(18).setCellValue(namecheck);
                    

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            
            try {
            	// for store the multiple we can use the time to store the multiple files
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Swiggy_OutputData" + timestamp + ".xlsx";
                
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
              //  driver.quit();
            }
        }
    }
}




//product name

/*      try {
          WebElement brandName = productElementInside.findElement(By.xpath(".//div[@class = 'sc-aXZVg gXaIUH _1v3Kq']"));
          WebElement nameOfProduct = productElementInside.findElement(By.xpath(".//div[@class = 'sc-aXZVg gnOsqr _AHZN']"));

           name1 = brandName.getText();
           name2 = nameOfProduct.getText();

           fullProductName = name1 + " " + name2;   
          // System.out.println(fullProductName);

      	WebElement brandName = productElementInside.findElement(By.xpath(" .//div[@class= '_7I7zN']//div[@class='sc-gEvEer jlNTHb _AHZN']"));
      	
      	name1 = brandName.getText();
      	
      	fullProductName = name1;
      	
      	
      	
      } catch (NoSuchElementException e) {
      	
      	for(int y=0;y<500;y++) {
      		
      		driver.navigate().refresh();
      		
      		Thread.sleep(2000);
      		
      		driver.navigate().refresh();
      		
      		Thread.sleep(3000);
          // Handle stale element reference exception by re-finding the element
      		
          WebElement brandName = productElementInside.findElement(By.xpath(".//div[@class='_1Cxye']/div[@class='hgjRZ']/div[2]"));
          WebElement nameOfProduct = productElementInside.findElement(By.xpath(" .//div[@class = 'sc-gEvEer jlNTHb _AHZN']"));

           name1 = brandName.getText();
           System.out.println(name1);
          name2 = nameOfProduct.getText();

           fullProductName = name1 + " " + name2;
      }
      	//break;
          // System.out.println(fullProductName);
      } catch (StaleElementReferenceException e) {
      	WebElement brandName = productElementInside.findElement(By.xpath(".//div[@class='_1Cxye']/div[@class='hgjRZ']/div[2]"));
      	 name1 = brandName.getText();
      	  System.out.println(name1);
          e.printStackTrace();
      }

      
      
      
      try {
          WebElement mrp = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _2XPBo _1QyO8']"));
          String originalMrp1 = mrp.getText();
          originalMrp = originalMrp1.replace("₹", "");
      } catch (StaleElementReferenceException e) {
          // Handle stale element reference exception by re-finding the element
      	
      	 WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
      	 
      	 WebElement mrp = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _2XPBo _1QyO8']")));
      //    WebElement mrp = productElement.findElement(By.xpath(".//div[@class='sc-aXZVg gXaIUH _2XPBo _1QyO8']"));
          String originalMrp1 = mrp.getText();
          originalMrp = originalMrp1.replace("₹", "");
      } catch (Exception e) {
          e.printStackTrace();
      }

      
      
      
      try {
          WebElement sp = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer jlNTHb _2XPBo']"));
          String originalSp = sp.getText();
          spValue = originalSp.replace("₹", "");
      } catch (StaleElementReferenceException e) {
          // Handle stale element reference exception by re-finding the element
          WebElement sp = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer jlNTHb _2XPBo']"));
          String originalSp = sp.getText();
          spValue = originalSp.replace("₹", "");
      } catch (Exception e) {
          e.printStackTrace();
      }

      
      
      
      
      try {
          WebElement uomProduct = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _1TwvP']"));
          originalUom = uomProduct.getText();
      } catch (StaleElementReferenceException e) {
          // Handle stale element reference exception by re-finding the element
          WebElement uomProduct = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _1TwvP']"));
          originalUom = uomProduct.getText();
      } catch (Exception e) {
          e.printStackTrace();
      }     
*/