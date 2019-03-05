package core;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.microsoft.sqlserver.jdbc.SQLServerException;
import com.mongodb.DB;

import utilities.DBManager;
import utilities.GetCaptcha;

public class Getdata {
	public WebDriver driver;
	public WebDriverWait wait;
	String baseurl = "https://www.justdial.com/";
	
	String city = "Delhi";
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

	DBManager dbm = new DBManager();
	
	DB db = null;

	String tableName = "Automation.JustdialDoctor";
	
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
		
		dbm.DBConnection(LiveDB_Path, Liveusename, Livepassword);
		}catch(Exception e){
			
		}
	}
	
	@Test
	public void getdata() throws Exception{
		
		driver.get("https://www.justdial.com/Delhi/Doctors");
		driver.navigate().refresh();
		driver.manage().window().maximize();
		
		  
		  for (int pageno =1;pageno<=5;pageno++){
			  
			  
			  String remarks = Integer.toString(pageno);
			  
			  baseurl = "https://www.justdial.com/Delhi/Doctors/nct-10892680/page-"+pageno;
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
	        System.out.println("The number of doctors on page is " + listcount);
	        
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
						DocData.put("Doctor Name : ", name);
						
						WebElement ele = pom.Homepage.phonenumber(driver);					
						
						File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
						String capText = GetCaptcha.captchatext(ele, driver, screenshot, 0, 0, 28,0, basePath);
						
						phonenumber = capText;
		
						Thread.sleep(4000);
						System.out.println("The number is " + phonenumber);
						DocData.put("Phone number is ", phonenumber);
						//address = pom.Homepage.address(driver).getText();
						DocData.put("Address: ", address);
						//time = pom.Homepage.timing(driver).getText();
						DocData.put("Doctor Timing : ", time);	
						try{
							driver.findElement(By.xpath("//*[@id='comp-contact']/li[2]/span[2]/a[3]")).click();
						}catch(Exception e){
							
						}
						//specialization = pom.Homepage.specialization(driver).getText();
						System.out.println("specialization " +specialization);
						DocData.put("Specialization : ", specialization);
						DocData.put("City ", city);
						try{
						about=pom.Homepage.about(driver).getText();
						}catch(Exception e){
							
						}
						
						List <WebElement> moredetails = driver.findElements(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div"));
						int morecount = moredetails.size();
						for (int j=4;j<=morecount;j++){
							
							
						String moretext = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div["+j+"]")).getText();
						if(moretext.contains("Listed")){
							//alsolisted=moretext;
							DocData.put("Also listed for: ", moretext);
						}else if(moretext.contains("Payment")){
							
							//payment=moretext;
							DocData.put("Payment mode : ", moretext);
						}else if(moretext.contains("Year")){
							
							//establishment=moretext;
							DocData.put("Year of Establishment : ", moretext);
						}
						
						
						}
					} catch (InterruptedException e) {
						// TODO Auto-generated catch block
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
	        
	}
	

