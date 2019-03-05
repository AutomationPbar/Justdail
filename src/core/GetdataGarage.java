package core;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.microsoft.sqlserver.jdbc.SQLServerException;
import com.mongodb.DB;
import utilities.DBManager;
import utilities.GetCaptcha;

public class GetdataGarage {
	public WebDriver driver;
	public WebDriverWait wait;
	String baseurl = "https://www.justdial.com/";
	
	String city = "Gurgaon";
	String name ="";
	String basepath = "C:\\eclipse\\justdial";
	String basePath = "C:\\eclipse\\justdial\\";
	
	String LiveDB_Path = "jdbc:sqlserver://10.0.10.42:1433;DatabaseName=PBCroma";
	private String Liveusename = "PBLIVE";
	private String Livepassword = "PB123Live";
	String phonenumber ="";
	String address ="";
	String time ="";
	String specialization = "";
	String alsolisted ="";
	String payment ="";
	String establishment ="";
	String about ="";
	String rating ="";
	String votes = "";
	String dayoff ="";
	String pincode = "";
	String alternate = "";

	DBManager dbm = new DBManager();
	
	DB db = null;

	String tableName = "Automation.JustdialGarage";
	
	String excelpath_update = "C:\\Excelfiles\\JustDial\\Justdialdataentry.xls";
	String sheetname = "Base Template";

	int rowCount;
	int excelrow = 1;
	
	XSSFSheet sheet;
	XSSFSheet modelsheet;
	XSSFRow row = null;
	XSSFWorkbook workbook;
	String resultdata[] = new String[15];
	
	String nodata ="No Data Found";
	
	@BeforeTest
	public void setup(){
		
		try{
		
		System.setProperty("webdriver.chrome.driver", "C://eclipse//chromedriver.exe");
		ChromeOptions options = new ChromeOptions();        
		options.addArguments("--disable-notifications");           
		new DesiredCapabilities();
		DesiredCapabilities caps = DesiredCapabilities.chrome();
		caps.setCapability(ChromeOptions.CAPABILITY, options);
		
		driver= new ChromeDriver(options);
		
		SimpleDateFormat formatter = new SimpleDateFormat("dd_MM_yyyy_HH_mm");
		Date datedd = new Date();
		System.out.println(formatter.format(datedd));
		String localDate11 = formatter.format(datedd).toString();
		excelpath_update = "C:\\Excelfiles\\JustDial\\Justdialdataentry" + localDate11 + ".xls";
		utilities.ExcelUtils.SetExcelFile(excelpath_update, sheetname);

		
		dbm.DBConnection(LiveDB_Path, Liveusename, Livepassword);
		}catch(Exception e){
			
		}
	}
	
	@Test
	public void getdata() throws Exception{
		
		driver.get("https://www.justdial.com/Gurgaon/Garages");
		driver.navigate().refresh();
		driver.manage().window().maximize();
		
		  
		  for (int pageno =1;pageno<=2;pageno++){
			  
			  
			  String remarks = Integer.toString(pageno);
			  
			  baseurl = "https://www.justdial.com/Gurgaon/Garages-in-Gurgaon/nct-10222543/page-"+pageno;
			  System.out.println("base url is " + baseurl);
			  driver.get(baseurl);
			  
			  JavascriptExecutor js = (JavascriptExecutor) driver;
			  js.executeScript("arguments[0].scrollIntoView();", pom.Homepage.dintfindbutton(driver));
			  try {
				  driver.manage().window().maximize();
				Thread.sleep(4000);
			} catch (InterruptedException e) {
				
			}
	        //This will scroll the web page till end.		
	        //js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
	
		
	        List <WebElement> doctorlist = driver.findElements(By.xpath("//*[@id='tab-5']/ul/li"));
	        
	        int listcount = doctorlist.size();
	        System.out.println("The number of garages on page is " + listcount);
	        
	        for (int i =0;i<listcount;i++){
	        	
	        	JSONObject DocData = new JSONObject();
	        	
	        	String docurl = doctorlist.get(i).getAttribute("data-href");
	        	
	        	System.out.println("the doctor url is " + docurl);
	        	
	        	((JavascriptExecutor) driver).executeScript("window.open()");

				ArrayList<String> handles = new ArrayList<String>(driver.getWindowHandles());

				System.out.println("Tabs :" + handles.size());

				if (handles.size() > 1) {

					driver.switchTo().window(handles.get(1));

					driver.get(docurl);
					try {
						Thread.sleep(3000);
						
						name = pom.Homepage.doctorname(driver).getText();
						DocData.put("Garage Name : ", name);
						try{
						rating = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[1]/div[2]/div/div/div/span/span[1]/span[1]")).getText();
						}catch(Exception e){
							
						}
						try{
							votes = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[1]/div[2]/div/div/div/span/span[2]/span[1]")).getText();
							}catch(Exception e){
								
							}
						
						WebElement ele = pom.Homepage.phonenumber(driver);					
						
						File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
						Thread.sleep(2000);
						String capText = GetCaptcha.captchatext(ele, driver, screenshot, 0, 0, 28,0, basePath);
					
						String[] ptext = capText.split(",");
						phonenumber = ptext[0];	
							
						
						Thread.sleep(4000);
						System.out.println("The number is " + phonenumber);
						DocData.put("Phone number ", phonenumber);
						address = pom.Homepage.address(driver).getText();
						DocData.put("Address: ", address);
						if(address.contains(city)){
							
							Pattern zipPattern = Pattern.compile("(\\d{6})");
							Matcher zipMatcher = zipPattern.matcher(address);
							if (zipMatcher.find()) {
							    pincode = zipMatcher.group(1);
							    System.out.println("The pincode is " + pincode);
							}
							
						}else {
							
							driver.close();
							driver.switchTo().window(handles.get(0));
							continue;
						}
						
						try{
						driver.findElement(By.xpath("//*[@id='vhall']")).click();
						}catch(Exception e){
							
						}
						Thread.sleep(1500);
						try{
						time = pom.Homepage.timing(driver).getText();
						}catch(Exception e){
							
						}
						
						DocData.put("Timing of all days : ", time);	
						
						for(int d =1;d<8;d++){
							
							String daytime = "";
							try{
							daytime = driver.findElement(By.xpath("//*[@id='statHr']/li["+d+"]/span[2]")).getText();
							time = driver.findElement(By.xpath("//*[@id='statHr']/li[1]/span[2]")).getText();
							
							if(daytime.contains("Closed")){
								
							dayoff = driver.findElement(By.xpath("//*[@id='statHr']/li["+d+"]/span[1]")).getText();
							dayoff = dayoff + "Closed";
							}else{
								dayoff = "All Days working";
							}
							}catch(Exception e){
								
							}
						}
						//DocData.put("Timing : ", time);	
						try{
							driver.findElement(By.xpath("//*[@id='comp-contact']/li[2]/span[2]/a[3]")).click();
						}catch(Exception e){
							
						}
						//specialization = pom.Homepage.specialization(driver).getText();
						System.out.println("specialization " +specialization);
						//DocData.put("Specialization : ", specialization);
						DocData.put("City ", city);
						try{
						about=pom.Homepage.about(driver).getText();
						}catch(Exception e){
							
						}
						
						try {

							WebElement ccName1 = driver.findElement(By.xpath("//*[@id='tab_1']/div[1]/div/img"));
							String logoSRC = ccName1.getAttribute("src");
							System.out.println("image url is " + logoSRC);

							URL imageURL = new URL(logoSRC);
							BufferedImage saveImage = ImageIO.read(imageURL);

							ImageIO.write(saveImage, "png", new File("C:\\ccimage\\" + name + ".png"));
							String imgpath = "C:\\ccimage\\" + name + ".png";

							String imgurl = utilities.Makeurl.geturl(imgpath);
							System.out.println("api url :" + imgurl);

							DocData.put("Image Url :", imgurl);

						} catch (Exception e) {
							e.printStackTrace();
						} 
						
						List <WebElement> moredetails = driver.findElements(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div"));
						int morecount = moredetails.size();
						for (int j=4;j<=morecount;j++){
							
							
						String moretext = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div["+j+"]")).getText();
						if(moretext.contains("Listed")){
							moretext =driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div[4]/ul")).getText();
							alsolisted=moretext;
							DocData.put("Also listed for: ", moretext);
							System.out.println(alsolisted);
							
						}else if(moretext.contains("Payment")){
							
							payment=moretext;
							DocData.put("Payment mode : ", moretext);
						}else if(moretext.contains("Year")){
							
							establishment=moretext;
							DocData.put("Year of Establishment : ", moretext);
						}
						
						
						}
					} catch (InterruptedException e) {
						
						//e.printStackTrace();
					}
					
					String jsondata = DocData.toString();
					System.out.println("Json data : " + jsondata);
					
					try {

						System.out.println("Going to insert data in table");
						city=Sanitize(city);
						name=Sanitize(name);
						specialization=Sanitize(specialization);
						payment=Sanitize(payment);
						phonenumber=Sanitize(phonenumber);
						address=Sanitize(address);
						time=Sanitize(time);
						alsolisted=Sanitize(alsolisted);
						establishment=Sanitize(establishment);
						about=Sanitize(about);
						docurl=Sanitize(docurl);
						
						
						resultdata[0] = name;
						resultdata[1] = city;
						resultdata[2] = address;
						resultdata[3] = phonenumber;
						resultdata[4] = pincode;
						resultdata[5] =dayoff;
						resultdata[6] =time;
						resultdata[7] = alsolisted;
						resultdata[8] = "0";
						resultdata[9] = "0";
						resultdata[10] = phonenumber;
						resultdata[11] = "Justdial";
						resultdata[12] = rating;
						resultdata[13] = votes;
						resultdata[14] = alternate;
						

						SetCellData1(excelpath_update, sheetname, resultdata, excelrow);
						excelrow++;
						
						dbm.SetPractoLabData1(city, name, jsondata, remarks,tableName, specialization,payment,phonenumber,address,time,alsolisted,establishment,about,docurl);

					} catch (SQLServerException e) {

						dbm.DBConnection(LiveDB_Path, Liveusename, Livepassword);

						 e.printStackTrace();

					} catch (Exception e) {

						

						 e.printStackTrace();

					}
				}
				
				driver.close();
				driver.switchTo().window(handles.get(0));
			}
	        	
	        }
		  
	}
	
public String Sanitize(String inp){
		
		String output = "";
		
		output = inp.replace("'", "");
		
		return output;
		
	}



public static void SetInputData(String filePath, String sheetName, int row, int col, String data) throws Exception {

	FileInputStream fis = new FileInputStream(filePath);
	XSSFWorkbook workbook = new XSSFWorkbook(fis);
	XSSFSheet inputSheet = workbook.getSheetAt(0);

	// Retrieve the row and check for null
	XSSFRow row0 = (XSSFRow) inputSheet.getRow(row);
	Cell cell = null;
	if (row0 == null) {
		row0 = (XSSFRow) inputSheet.createRow(row);
	}
	// Update the value of cell
	cell = row0.getCell(col);
	if (cell == null) {
		cell = row0.createCell(col);
	}
	cell.setCellValue(data);

	try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
		workbook.write(fileOut);
		fileOut.close();
	} catch (Exception e) {
		System.out.println(e);
	}

	workbook.close();
}

public static void SetCellData1(String filePath, String sheetName, String[] result, int row) throws Exception {

	FileInputStream ExcelFile = new FileInputStream(filePath);

	HSSFWorkbook wb = new HSSFWorkbook(ExcelFile);

	Sheet resultSheet = wb.getSheet(sheetName);

	System.out.println("Row Passed : " + row);

	if (row == 1) {
		Row row0 = resultSheet.createRow(0);

		row0.createCell(0).setCellValue("S.No.");
		row0.createCell(1).setCellValue("Garage");
		row0.createCell(2).setCellValue("City");
		row0.createCell(3).setCellValue("Address");
		row0.createCell(4).setCellValue("Phone Number");
		row0.createCell(5).setCellValue("Pin Code");
		row0.createCell(6).setCellValue("Days of Opertaion");
		row0.createCell(7).setCellValue("Timing");
		row0.createCell(8).setCellValue("Make");
		row0.createCell(9).setCellValue("Latitude");
		row0.createCell(10).setCellValue("Longitude");
		row0.createCell(11).setCellValue("Login Mobile No");
		row0.createCell(12).setCellValue("Source");
		row0.createCell(13).setCellValue("Rating");
		row0.createCell(14).setCellValue("Votes");

	}
	Row row2 = resultSheet.createRow(row);
	row2.createCell(0).setCellValue(row);
	System.out.println("Row Created :" + (row));
	
	for (int i = 0; i < result.length; i++) {

		row2.createCell(i + 1).setCellValue(result[i]);

	}

	try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
		wb.write(fileOut);
		fileOut.close();
	} catch (Exception e) {
		System.out.println(e);
	}
	wb.close();

}
	        
	}
	

