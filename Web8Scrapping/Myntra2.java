package Web8Scrapping;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.remote.RemoteWebDriver;

import CommonUtility.ProductDetailsScrapping2;
import CommonUtility.UserUtility;

public class Myntra2 {
	
	
	public static void main(String[] args) {
		
	
		RemoteWebDriver driver = null; 
		int rowNum = 1;
		
		int ExceptionCount = 0;
		
		UserUtility userUtility = new UserUtility(UserUtility.InputDataFilePath2,"Sheet2", "Locators");		
		while (rowNum < userUtility.getNoOfRows(userUtility.InputDataSheet)) {
			try {
				ProductDetailsScrapping2 pds = new ProductDetailsScrapping2();
				driver = userUtility.launchBrowser(driver, false, false, false);
				
				
			//	Thread.sleep(5000);
				
				for (int i = 1; i < userUtility.getNoOfRows(userUtility.InputDataSheet); i++) {
					//driver.manage().deleteAllCookies();
					Row row = userUtility.InputDataSheet.getRow(rowNum);

					String InputBrandName = row.getCell(0).getStringCellValue().trim();
					String InputPid = row.getCell(1).getStringCellValue().trim();
					String InputCity = row.getCell(2).getStringCellValue().trim();
					String InputTitle = row.getCell(3).getStringCellValue().trim();
					String InputSize = row.getCell(4).getStringCellValue().trim();
					String OldProductCode = row.getCell(5).getStringCellValue().trim();
					String OldURL = row.getCell(6).getStringCellValue().trim();
					String OldName = row.getCell(7).getStringCellValue().trim();
					String OldMRP = row.getCell(8).getStringCellValue().trim();
					String OldSP = row.getCell(9).getStringCellValue().trim();
					String OldUOM = row.getCell(10).getStringCellValue().trim();
					String OldMultiplier = row.getCell(11).getStringCellValue().trim();
					String OldAvailability = row.getCell(12).getStringCellValue().trim();
					String OldOffer =row.getCell(13).getStringCellValue().trim();
					
					if(InputBrandName.contains("Nykaa")){
						driver.manage().deleteAllCookies();
					}			
					
					List<String> productDetails = pds.productDetails(driver, userUtility.locatorsMap.get(InputBrandName),
							"NA", false, OldURL, InputPid, InputBrandName);
					
					String NewProductCode = productDetails.get(0).trim();
					String NewURL = productDetails.get(1).trim();
					String NewName = productDetails.get(2).trim();
					String NewMRP = productDetails.get(3).trim();
					String NewSP = productDetails.get(4).trim();
					String NewUOM = productDetails.get(5).trim();
					String NewMultiplier = productDetails.get(6).trim();
					String NewAvailability = productDetails.get(7).trim();
					String NewOffer = productDetails.get(8).trim();
					String NewManualInterventionFlag = productDetails.get(9).trim();
					String LogFile = productDetails.get(10).trim();

					String MRPPriceValidation = "NA";
					String sellingPriceValidation = "NA";
					String ProductCodeValidation = "NA";
					String OldNameCheck = "NA";
					String CrtOffer = "NA";
					
					if (!OldProductCode.equals("NA")) {
						
						if (NewURL.contains(OldProductCode)) {
							NewProductCode = OldProductCode;
							ProductCodeValidation = "OK";
						}
						else {
							NewManualInterventionFlag = "YES";
							LogFile = LogFile + "Mismatch in the product code or URL /---------";
						}
						
						if (OldName.equals(NewName)) {
							if (OldMRP.equals(NewMRP)) {
								NewUOM = OldUOM;
								NewMultiplier = OldMultiplier;
								NewAvailability = OldAvailability;
							}
							else {
								NewUOM = OldUOM;
								NewMultiplier = OldMultiplier;
								NewAvailability = OldAvailability;
								NewManualInterventionFlag = "YES";
								LogFile = LogFile + "Old MRP and New MRP is not matching /---------";
								
							}
						}

						else {
							NewUOM = OldUOM;
							NewMultiplier = OldMultiplier;
							NewAvailability = OldAvailability;
							NewManualInterventionFlag = "YES";
							LogFile = LogFile + "Old Name and New Name is not matching /---------";
						}
						
						MRPPriceValidation = OldMRP;
						sellingPriceValidation = OldSP;
						
						OldNameCheck = OldName;
					}
					
					
					
//					try {
//
//							
//							if (!(NewSP.equals(" ") || NewSP.equals("") || NewSP.equals("NA"))) {
//								if (userUtility.checkDifferenceBTWNumbers(OldSP, NewSP, 20)) {
//									NewManualInterventionFlag = "YES";
//									LogFile = LogFile + "There is difference of more than 20% in SP /---------";
//								}
								
//								if (userUtility.checkDifferenceBTWNumbers(OldMRP, NewMRP, 3)) {
//									NewMRP = "Need to check";
//									MRPPriceValidation = "OK";
//								} 
//							}
//
//						}
//					 catch (Exception e) {
//						e.printStackTrace();
//					}
						if(NewMRP == NewSP) {
							NewOffer = "NA";
						}else if(OldOffer.equals(NewOffer)) {
							NewOffer = OldOffer;
						}
						
						if(NewOffer.contains("OFF") && NewOffer != " ") {
							CrtOffer = NewOffer.replaceAll("[()]","").replaceAll("OFF", "Off");
						}
					/*	//String getCode = NewOffer;
						else if(NewOffer.contains("-") && NewOffer != " ") {
							//CrtOffer = NewOffer.replaceAll("-","").replaceAll("%", "% Off");
							 CrtOffer = NewOffer.replace("-", "").trim() + " Off";
						}
						*/
						else {
							CrtOffer = NewOffer;
						}
						
						// Out of Stock check
						
						
						if(NewURL.contains("NA")){
							String result = "NA";
						}	
						
						int result = 1;
						try {
						String xpathForMyntra = "//div[@class='pdp-add-to-bag pdp-button pdp-flex pdp-center pdp-out-of-stock ']";

						
						String xpathForNykaa = "//span[@class='css-1neql7s']";
						
						boolean isElementPresent = !driver.findElements(By.xpath(xpathForMyntra)).isEmpty();   //is empty = true

				        if (!isElementPresent) {
				            isElementPresent = !driver.findElements(By.xpath(xpathForNykaa)).isEmpty();
				        }
				        
				        //true =0,false =1 

				        result = isElementPresent ? 0 : 1;
				        
				        System.out.println(result);
						}
						catch(Exception e) {
							System.out.println(e.getMessage());
						}
						
						//int stock = result;
						String NewAvailability1 = String.valueOf(result);
						

					List<String> finalList = new ArrayList<String>();

					finalList.add(InputPid);
					finalList.add(InputBrandName);
					finalList.add(InputTitle);
					finalList.add(InputSize);
					finalList.add(NewProductCode);
					finalList.add(NewURL);
					finalList.add(NewName);
					finalList.add(NewMRP);
					finalList.add(NewSP);
					finalList.add(NewUOM);
					finalList.add(NewMultiplier);
					finalList.add(NewAvailability1);
					finalList.add(CrtOffer);
					finalList.add(NewManualInterventionFlag);
					finalList.add(LogFile);
					finalList.add(MRPPriceValidation);
					finalList.add(sellingPriceValidation);
					finalList.add(ProductCodeValidation);
					finalList.add(OldNameCheck);
					
					System.out.println(finalList);

					if (rowNum == 1) {
						List<String> header = userUtility.header("");
						userUtility.writeIntoSheet(userUtility.OutputSheet, header);
					}

					userUtility.writeIntoSheet(userUtility.OutputSheet, finalList);

					rowNum++;
				}
				userUtility.writeIntoExcel(userUtility.OutputWorkBook, userUtility.OutputDataFilePath);
				System.out.println("Scrapping is completed");
			} 
			
			catch (Exception e) {
				System.out.println(e.getMessage());
				userUtility.writeIntoExcel(userUtility.OutputWorkBook, userUtility.OutputDataFilePath);
				try {
					driver.quit();
				} catch (Exception e1) {
					e1.printStackTrace();
				}
				
				ExceptionCount++;
				
				if (ExceptionCount>5){
					break;
				}
			}
		}
		
		try {
			driver.quit();
		} catch (Exception e1) {
			e1.printStackTrace();
		}
		
//		System.out.println(userUtility.locatorsMap.get("Myntra"));
//		System.out.println(userUtility.getExcelDataMap(userUtility.locatorsMap, "Myntra"));
//		System.out.println(userUtility.getExcelData(userUtility.locatorsMap, "Myntra", "Product Name"));
		
		
	}
	
	
	
	

}
