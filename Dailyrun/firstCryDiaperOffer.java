package Dailyrun;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.LinkedHashMap;

public class firstCryDiaperOffer {

    public static void main(String[] args) throws Exception{
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        
         int count = 0;
          String spValue = "";
          String finalSp = "";
          String newName = null;
          String mrpValue = null;
          String Offer2= "NA";
          String Offer1 = "NA";
          String discount1= "NA";
          String discount2= "NA";
          String discount3= "NA";
          String sp1= "NA";
          String sp2= "NA";
          String sp3= "NA";
          String sp4= "NA";
          String sp5= "NA";
          String NewAvailability1 = " ";
          
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\firstcryDiaDel.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("FirstCry2");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

	            List<String> inputPid = new ArrayList<>(),InputCity = new ArrayList<>(),InputName = new ArrayList<>(),InputSize = new ArrayList<>(),NewProductCode = new ArrayList<>(),
	            		uRL = new ArrayList<>(),UOM = new ArrayList<>(),Mulitiplier = new ArrayList<>(),Availability = new ArrayList<>();
	            
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
                    
                    inputPid.add(id);
                    InputCity.add(city);
                    InputName.add(name);
                    InputSize.add(size);
                    NewProductCode.add(productCode);
                    uRL.add(url);
                    UOM.add(uom);
                    Mulitiplier.add(mulitiplier);
                    Availability.add(availability);
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
            headerRow.createCell(10).setCellValue("MULTIPLIER");
            headerRow.createCell(11).setCellValue("AVAILABILITY");
            headerRow.createCell(12).setCellValue("OFFER 1");
            headerRow.createCell(13).setCellValue("SP 1");
            headerRow.createCell(14).setCellValue("OFFER 2");
            headerRow.createCell(15).setCellValue("SP 2");
            headerRow.createCell(16).setCellValue("DISCOUNT 1");
            headerRow.createCell(17).setCellValue("SP 3");
            headerRow.createCell(18).setCellValue("DISCOUNT 2");
            headerRow.createCell(19).setCellValue("SP 4");
            headerRow.createCell(20).setCellValue("DISCOUNT 3");
            headerRow.createCell(21).setCellValue("SP 5");
            
            int rowIndex = 1;

            int headercount = 0;
            
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
                          resultRow.createCell(13).setCellValue("NA");
                          resultRow.createCell(14).setCellValue("NA");
                          resultRow.createCell(15).setCellValue("NA");
                          resultRow.createCell(16).setCellValue("NA");
                          resultRow.createCell(17).setCellValue("NA");
                          resultRow.createCell(18).setCellValue("NA");
                          resultRow.createCell(19).setCellValue("NA");
                          resultRow.createCell(20).setCellValue("NA");
                          resultRow.createCell(21).setCellValue("NA");
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	 
                     if(i == 0 ) {
                    	 driver.get("https://www.firstcry.com");
                         driver.manage().window().maximize();
                         
                      Thread.sleep(1000);
                      WebElement locationSet= driver.findElement(By.xpath("//*[@id=\"geoLocation\"]/span/div[1]/span"));
                      locationSet.click();
                      Thread.sleep(1000);
                      
                      WebElement setLocation = driver.findElement(By.xpath("//*[@id=\"pincodetext\"]/div/sapn"));
                      setLocation.click();
                      
                      WebElement clickLocation = driver.findElement(By.xpath("//*[@id=\"nonlpincode\"]"));
                      clickLocation.clear();
                      clickLocation.sendKeys("110015");
                      
                      WebElement clickApply = driver.findElement(By.xpath("//*[@id=\"epincode\"]/div"));
                      clickApply.click();
                      
                      Thread.sleep(5000);
                      WebElement reg=driver.findElement(By.xpath("/html/body/div[1]/div[5]/div/div[3]/ul/li[7]"));
                      reg.click();
                      Thread.sleep(5000);
                      
                      WebElement regmail=driver.findElement(By.xpath("//*[@id=\"lemail\"]"));
                      regmail.click();
                      Thread.sleep(5000);
                      regmail.sendKeys("blktpoc2000@gmail.com");//blktpoc2000@gmail.com
                     //  blinkitproofofconcept@gmail.com
                      
                      WebElement conmail= driver.findElement(By.xpath("//*[@id=\"login\"]/div/div[3]/span"));
                      conmail.click();
                      
                      Thread.sleep(30000);
                      } 
                     
                	  driver.get(url);
                	  
                	  String addToCartButtonXPath1 = "/html/body/div[5]/div[2]/div/div[2]/div[7]/div/div[2]/div[2]/div[1]/div/span[1]/span";
                	  String addToCartButtonXPath2 = "/html/body/div[5]/div/div[2]/div[1]/div[1]/div[2]/div[2]/div[1]/div/span[1]/span";
                	  WebElement addToCartButton = null;

                	  try {
                	      addToCartButton = driver.findElement(By.xpath(addToCartButtonXPath1));
                	  } catch (NoSuchElementException e1) {
                	      try {
                	          addToCartButton = driver.findElement(By.xpath(addToCartButtonXPath2));
                	      } catch (NoSuchElementException e2) {
                	          System.out.println("Add to cart button not found.");
                	      }
                	  }
                	  
                	  if (addToCartButton != null && addToCartButton.isEnabled() && addToCartButton.isDisplayed()) {
                		  System.out.println("Add to Cart button is present on the page.");
                	  
                      try {
                      	
                      WebElement nameElement = driver.findElement(By.id("prod_name"));
                      newName = nameElement.getText();
                      System.out.println(newName);
                      }
                      
                      catch(org.openqa.selenium.NoSuchElementException e) {
                    	  try {
                      	
                      	WebElement nameElement = driver.findElement(By.xpath("//div[@class = 'prod-info-wrap']//following::p[1]"));
                      	newName = nameElement.getText();
                          System.out.println(newName);
                    	  }
                    	  catch(Exception h) {
                    		  WebElement nameElement = driver.findElement(By.xpath("//div[@class='right-contr']//div[@class='prod-info-wrap']//p[@class='prod-name R20_21']"));
                            	newName = nameElement.getText();
                                System.out.println(newName);
                    	  }
                      	
                      }
                      
                      System.out.println("headercount = " + headercount);
                      
                      headercount++;
                      
                      int Availability0 = 1;
                      NewAvailability1  = Integer.toString(Availability0);
                      
                      try {
                      WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));
                      mrpValue = mrp.getText();
                      System.out.println(mrpValue);
                      
                      }
                      
                      catch(org.openqa.selenium.NoSuchElementException e){
                      		try {
                          WebElement mrp = driver.findElement(By.xpath("/html/body/div[5]/div/div[2]/div[2]/div[2]/div[2]/span[4]/span[3]"));
                          mrpValue = mrp.getText();                            
                          System.out.println(mrpValue);
                      		}
                      		catch(Exception o) {
                      			 WebElement mrp = driver.findElement(By.xpath("//span[@class='pos-rel2stat new-mrp-wrap']//span[@class='pmr R20_75 pos-rel2stat']"));
                                 mrpValue = mrp.getText();                            
                                 System.out.println(mrpValue);
                      		}
                      	
                      }
                      Thread.sleep(2000);
                      List<WebElement> divElements = driver.findElements(By.xpath("//div[@class='cpn_offrs_disc_section']//div[@class='swiper-slide']"));
                      Map<String, List<String>> keyValueMap = new LinkedHashMap<>(); // Use LinkedHashMap to maintain insertion order
                      List<String> last7CharsList = new ArrayList<>();
                      int divCount = divElements.size();
                      System.out.println("Number of slides: " + divCount);

                      WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10)); 
                      
                      String last7Chars = "";  
                      String code = "";  

                      for (int j = 1; j <= divCount; j++) {
                          try {
                              String xpath = "(//div[@class='cpn_offrs_disc_section']//div[@class='swiper-slide'])[" + j + "]";

                              WebElement offerElement = wait.until(ExpectedConditions.visibilityOfElementLocated(
                                      By.xpath(xpath + "//div[@class='club-txt-logo']/parent::div/span")
                              ));
                              String offer = offerElement.getText();
                              
                              if (offer.contains("Extra") || offer.contains("FLAT") || offer.contains("extra") || offer.contains("Flat")) {
                                  String discountText = offer;

                                  if (discountText.contains("Non Club") || discountText.contains("FLAT") || discountText.contains("Extra") || discountText.contains("extra") || offer.contains("Flat")) {
                                	  
                     
                                      String[] parts = discountText.split(" ");
                                      int indexOfFlat = -1;
                                      int indexOfExtra = -1;

                                      for (int b = 0; b < parts.length; b++) {
                                          if (parts[b].equalsIgnoreCase("Flat")) {
                                              indexOfFlat = b; 
                                          }
                                          if (parts[b].equalsIgnoreCase("Extra")) {
                                              indexOfExtra = b;
                                          }
                                      }

                                      if (indexOfFlat != -1 && indexOfFlat + 1 < parts.length) {
                                          String flatDiscountValue = parts[indexOfFlat + 1]; // This should be "35%"
                                          System.out.println("Flat Discount Value: " + flatDiscountValue);
                                      }

                                      if (indexOfExtra != -1 && indexOfExtra + 1 < parts.length) {
                                          String extraDiscountValue = parts[indexOfExtra + 1]; // This should be "15%"
                                          System.out.println("Extra Discount Value: " + extraDiscountValue);
                                      }

                                      if (discountText.contains("*")) {
//                                          int starIndex = discountText.indexOf("*");
//                                          String beforeStar = discountText.substring(0, starIndex);
//                                          last7Chars = beforeStar.length() > 7 ? beforeStar.substring(beforeStar.length() - 7) : beforeStar;
//
//                                          System.out.println("Last 7 Characters Before '*': " + last7Chars);
                                    	  
                                    	  int lastStarIndex = discountText.lastIndexOf("*");
                                          if (lastStarIndex != -1) {
                                              // Extract the portion of the string before the last asterisk
                                              String beforeLastStar = discountText.substring(0, lastStarIndex);
                                              // Get the last 7 characters before the last asterisk
                                              last7Chars = beforeLastStar.length() > 7 ? beforeLastStar.substring(beforeLastStar.length() - 7) : beforeLastStar;

                                              System.out.println("Last 7 Characters Before Last '*': " + last7Chars);
                                          }

                                          String codeXpath = xpath + "//ancestor::div[@class='save_cpn_header']/following-sibling::div//div[contains(@class,'cpn_code_box')]";
                                          WebElement codeElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(codeXpath)));
                                          code = codeElement.getText(); 
                                          System.out.println("Coupon Code: " + code);

                                          keyValueMap.computeIfAbsent(discountText, k -> new ArrayList<>()).add(code);
                                          last7CharsList.add(last7Chars);
                                      } else {
                                          System.out.println("No asterisk found in the text.");
                                      }
                                  } else {
                                      System.out.println("The text does not contain 'Non Club'.");
                                  }
                              }
                          } catch (Exception e) {
                              e.printStackTrace(); 
                          }
                      }

                      System.out.println("Final key-value pairs in the map:");
                      for (Map.Entry<String, List<String>> entry : keyValueMap.entrySet()) {
                          String offerPercentage = entry.getKey(); 
                          List<String> couponCodes = entry.getValue(); 
                          System.out.println("Offer Percentage: " + offerPercentage + ", Coupon Codes: " + couponCodes);
                      }

                      try {
                          Thread.sleep(1000);
                          driver.findElement(By.xpath("(//span[@class='step1 M16_white'])[1]//span")).click();
                          Thread.sleep(1000);
                          driver.findElement(By.xpath("(//span[@class='step2 M16_white'])[1]")).click();
                      } catch (NoSuchElementException e) {
                      }

                      Thread.sleep(2000);
                      
                      String rateValue = "";
                      WebDriverWait wait1 = new WebDriverWait(driver,Duration.ofSeconds(10));
                      Thread.sleep(4000);
                    	 
                    	  wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[4]/div[4]/div[10]/div[4]/span[2]")));

                      WebElement rate = driver.findElement(By.xpath("/html/body/form/section[1]/section[4]/div[4]/div[10]/div[4]/span[2]"));
                     
                      rateValue = rate.getText();
                      double amount1 = Double.parseDouble(rateValue) / 100.0;
                      String formattedAmount1 = String.format("%.2f", amount1);
          	        System.out.println("Final value for coupon code : " + formattedAmount1);
                      System.out.println(formattedAmount1);
                      finalSp = formattedAmount1;
                      
                      BlinkitId screenshot = new BlinkitId();
                	  try {
            				screenshot.screenshot(driver, "Firstcry", id);
            			} catch (Exception e) {
            				e.fillInStackTrace();
            			
            			}
                      
                      Thread.sleep(3000);

                      List<String> couponCodesList = new ArrayList<>();
                      List<String> offersList = new ArrayList<>(keyValueMap.keySet()); 
                      for (List<String> codes : keyValueMap.values()) {
                          couponCodesList.addAll(codes); 
                      }

                      Actions actions = new Actions(driver);
                      for (int p = 0; p < offersList.size(); p++) {
                          String couponCode = couponCodesList.get(p);
                          String offerText = last7CharsList.get(p);
                          System.out.println("Applying Offer: " + offerText + ", Coupon Code: " + couponCode);

                          WebElement coupon = driver.findElement(By.xpath("//div[@class='cupn_cod']//div[@class='input_field coup_inputfied div_input']//input "));
                          Thread.sleep(5000);
                          actions.moveToElement(coupon).click().perform();

                          Thread.sleep(500);
                          coupon.clear();
                          Thread.sleep(500);
                          coupon.sendKeys(couponCode);
                          Thread.sleep(1000);

                          wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/section[1]/section[4]/div[4]/div[8]/div[2]/div[3]/div/span[2]")));
                          WebElement applyClick = driver.findElement(By.xpath("/html/body/form/section[1]/section[4]/div[4]/div[8]/div[2]/div[3]/div/span[2]"));
                          actions.moveToElement(applyClick).click().perform();
                          Thread.sleep(6000);

                          WebElement rate1 = driver.findElement(By.xpath("/html/body/form/section[1]/section[4]/div[4]/div[10]/div[4]/span[2]"));
                          String rateValue1 = rate1.getText();
                          double amount = Double.parseDouble(rateValue1) / 100.0;

                          String formattedAmount = String.format("%.2f", amount);
                          System.out.println("Rate value for coupon code " + couponCode + ": " + formattedAmount);

                          Thread.sleep(2000);
                          String xpathExpression = "//div[@class='input_field coup_inputfied div_input']//p[@class='J12M_42 cl_e5 errmsg err1']";

                          try {
                              WebElement invalidCouponElement = driver.findElement(By.xpath(xpathExpression));
                              System.out.println("Invalid coupon message displayed.");
                          } catch (Exception g) {
                              WebElement elements = driver.findElement(By.id("coponapply"));
                              elements.click();
                              System.out.println("'Invalid coupon' text not found on the webpage.");
                          }
                          String extractOffer = offerText;

                          if ("8% Off".equals(extractOffer) || "7% Off".equals(extractOffer)) {
                              String offerText1 = "5% Off";
                              offerText = offerText1;
                          }

                          if (offerText == null || couponCode == null || offerText.isEmpty() || couponCode.isEmpty()) {
                              switch (p) {
                                  case 0:
                                      discount1 = "NA";
                                      sp3 = "NA";
                                      break;
                                  case 1:
                                      discount2 = "NA";
                                      sp4 = "NA";
                                      break;
                                  case 2:
                                      discount3 = "NA";
                                      sp5 = "NA";
                                      break;
                                  case 3:
                                      Offer1 = "NA";
                                      sp1 = "NA";
                                      break;
                                  case 4:
                                      Offer2 = "NA";
                                      sp2 = "NA";
                                      break;
                                  default:
                                      break;
                              }
                          } else {
                              switch (p) {
                                  case 0:
                                	  Offer1 = offerText;
                                      sp1 = formattedAmount;
                                      break;
                                  case 1:
                                	  Offer2 = offerText;
                                      sp2 = formattedAmount;
                                      break;
                                  case 2:
                                      discount1 = offerText;
                                      sp3 = formattedAmount;
                                      break;
                                  case 3:
                                	  discount2 = offerText;
                                      sp4 = formattedAmount;
                                      break;
                                  case 4:
                                	  discount3 = offerText;
                                      sp5 = formattedAmount;
                                      break;
                                  default:
                                      break;
                              }
                          }
                          
                          System.out.println("Coupon after applying coupon " + (p + 1) + ": " + couponCode);
                          System.out.println("SP after applying coupon " + (p + 1) + ": " + formattedAmount);
                      }
                           	
                      try {
                      	Thread.sleep(2000); 
                       WebElement remove = driver.findElement(By.xpath("//div[@class='short_prod newshort']//div[@class='new-shortone shortcomm']"));

                       remove.click();
                      } catch (Exception e) {
                         
                      	Thread.sleep(4000); 
                       WebElement remove = driver.findElement(By.xpath("//div[@class='short_prod newshort']//div[@class='new-shortone shortcomm']"));
                       remove.click();

                          e.printStackTrace();
                          Thread.sleep(2000); 
                      }
                	          
                	  }
                	  else {
                		  
                		  System.out.println("Add to Cart button is NOT present on the page.");
                		  
                		  Thread.sleep(2000);
                		  
                		  int Availability1 = 1;
                          
                		  try {
                            	
                              WebElement nameElement = driver.findElement(By.id("prod_name"));
                              newName = nameElement.getText();
                              System.out.println(newName);
                              }
                              
                              catch(org.openqa.selenium.NoSuchElementException e) {
                              	WebElement nameElement = driver.findElement(By.xpath("//div[@class = 'prod-info-wrap']//following::p[1]"));
                              	newName = nameElement.getText();
                                  System.out.println(newName);
                              	
                              }
                              
                              try {
                              WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));
                              mrpValue = mrp.getText();
                              System.out.println(mrpValue);
                              
                              }
                              
                              catch(org.openqa.selenium.NoSuchElementException e){
                              	
                                  WebElement mrp = driver.findElement(By.xpath("/html/body/div[5]/div/div[2]/div[2]/div[2]/div[2]/span[4]/span[3]"));
                                  mrpValue = mrp.getText();                            
                                  System.out.println(mrpValue);
                              	
                              }
                          Thread.sleep(500);
                          
                          WebElement sp = driver.findElement(By.xpath("//*[@id='prod_price']"));
                         spValue = sp.getText();
                         System.out.println("+++++++++++" + spValue);
                          finalSp = spValue;
                         
                          Availability1 = 0;
                          NewAvailability1  = Integer.toString(Availability1);
                          
                         count=1;
                         
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
                      resultRow.createCell(8).setCellValue(finalSp);
                      resultRow.createCell(9).setCellValue(uom);
                      resultRow.createCell(10).setCellValue(mulitiplier);
                      resultRow.createCell(11).setCellValue(NewAvailability1);
                      resultRow.createCell(12).setCellValue(Offer1);
                      resultRow.createCell(13).setCellValue(sp1);
                      resultRow.createCell(14).setCellValue(Offer2);
                      resultRow.createCell(15).setCellValue(sp2);
                      resultRow.createCell(16).setCellValue(discount1);
                      resultRow.createCell(17).setCellValue(sp3);
                      resultRow.createCell(18).setCellValue(discount2);
                      resultRow.createCell(19).setCellValue(sp4);
                      resultRow.createCell(20).setCellValue(discount2);
                      resultRow.createCell(21).setCellValue(sp5);
                      
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
                      resultRow.createCell(9).setCellValue("NA");
                      resultRow.createCell(10).setCellValue("NA");
                      resultRow.createCell(11).setCellValue("NA");
                      resultRow.createCell(12).setCellValue("NA");
                      resultRow.createCell(13).setCellValue("NA");
                      resultRow.createCell(14).setCellValue("NA");
                      resultRow.createCell(15).setCellValue("NA");
                      resultRow.createCell(16).setCellValue("NA");
                      resultRow.createCell(17).setCellValue("NA");
                      resultRow.createCell(18).setCellValue("NA");
                      resultRow.createCell(19).setCellValue("NA");
                      resultRow.createCell(20).setCellValue("NA");
                      resultRow.createCell(21).setCellValue("NA");
                      
                      System.out.println("Failed to extract data for URL: " + url);
                      
                  }
                Offer1 = "NA";
                sp1 = "NA";
                Offer2 = "NA";
                discount1 = "NA";
                discount2 = "NA";
                discount3 = "NA";
                sp2 = "NA";
                sp3 = "NA";
                sp4 = "NA";
                sp5	 = "NA";
              }
              
            try {
            	// for store the multiple we can use the time to store the multiple files
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Firstcry_Diapers_OutputData_SecondHalf" + timestamp + ".xlsx";
                
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
                  driver.quit();
              }
          }
        
        }
}