import org.openqa.selenium.By;
import org.openqa.selenium.By.ByClassName;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.webdriver.WebDriverBrowser;

public class DemoWebPage {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
WebDriver driver =new ChromeDriver();
driver.get("https://demowebshop.tricentis.com/");
driver.manage().window().maximize();
//driver.findElement(By.className("ico-register")).click();
driver.findElement(By.linkText("Register")).click();
driver.findElement(By.className("ico-login")).click();
String expectedText = "divyanburose@gmail.com";
driver.findElement(By.id("Email")).sendKeys(expectedText);	
driver.findElement(By.id("Password")).sendKeys("divyanbu");
driver.findElement(By.xpath("//input[@value=\"Log in\"]")).click();
String actualText = driver.findElement(By.className("account")).getText();
System.out.println(actualText);	
if(expectedText.equals(actualText)) {
System.out.println("Login success");
}else {
	System.out.println("Login unsuccess");
}
//backward

String ActualPage = driver.getTitle();
System.out.println("ActualPage:"+ActualPage);
driver.navigate().back();

//forward

String ExpectedPage = driver.getTitle();
System.out.println("ExpectedPage:"+ExpectedPage);
driver.navigate().forward();

if (ExpectedPage.equals("Demo Web Shop. Login")) {
	System.out.println("navigation success");
}
	
driver.navigate().refresh();	
	
	}

}
