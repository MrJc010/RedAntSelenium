import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
public class LoginPage {
	
	private WebDriver driver;
	private By uname = By.name("uname");
	private By psw = By.xpath("//input[@id='psw']");
	private By loginButton = By.xpath("//img[@id='subImg']");
	
	public LoginPage(WebDriver driver) {
		this.driver = driver;
	}
	
	public WebElement getUnameText(WebDriver driver) {
		return driver.findElement(uname);
	}
	
	public WebElement getPassWordText(WebDriver driver) {
		return driver.findElement(psw);
	}
	
	public WebElement getLoginButton(WebDriver driver) {
		return driver.findElement(loginButton);
	}
	

}
