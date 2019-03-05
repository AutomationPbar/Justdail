package pom;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;




public class Homepage {
	
	WebDriver driver;
	
	public static WebElement selectcity(WebDriver driver){
		
		WebElement selcity = driver.findElement(By.xpath("//*[@class='search-city mnsrchwpr']"));
		return selcity;
	}
	
public static WebElement dintfindbutton(WebDriver driver){
		
		WebElement db = driver.findElement(By.xpath("//*[@id='setbackfix']/span/a"));
		return db;
	}
	
	public static WebElement searchbox(WebDriver driver){
		
		WebElement sb = driver.findElement(By.className("search-text"));
		return sb;
	}
	
	public static WebElement searchbutton(WebDriver driver){
		
		WebElement sb = driver.findElement(By.id("srIconwpr"));
		return sb;
	}
	public static WebElement doctorname(WebDriver driver){
	
	WebElement dn = driver.findElement(By.className("fn"));
	return dn;
	}
	
	public static WebElement phonenumber(WebDriver driver){
		
		//*[@id="setbackfix"]/div[1]/div/div[1]/div[2]/div/ul/li[2]/p/span[2]/span/a
		//*[@id='comp-contact']/span[2]/a
		WebElement pn = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[1]/div[2]/div/ul/li[2]/p/span[2]/span/a"));
		return pn;
		}
	public static WebElement address(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='fulladdress']/span/span"));
		return pn;
		}
	
	public static WebElement specialization(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='comp-contact']/li[2]"));
		return pn;
		}
	public static WebElement viewall(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='vhall']"));
		return pn;
		}
	public static WebElement timing(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='mhd']"));
		return pn;
		}
	
	public static WebElement alsolisted(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div[4]"));
		return pn;
		}
	
	public static WebElement payment(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div[5]"));
		return pn;
		}
	public static WebElement year(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[1]/div[6]"));
		return pn;
		}
	
public static WebElement about(WebDriver driver){
		
		WebElement pn = driver.findElement(By.xpath("//*[@id='setbackfix']/div[1]/div/div[4]/div[2]/div[7]/span/p[6]"));
		return pn;
		}
}
