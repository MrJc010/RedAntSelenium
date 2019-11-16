import java.io.File;
import java.io.IOException;
import java.nio.file.DirectoryNotEmptyException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;

import javax.swing.JOptionPane;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class RunPageLogin {
	WebDriver driver = null;
	private String uname = "vietn";
	private String psw = "biz285";
	LoginPage loginPage = null;
	WebElement userName = null;
	WebElement password = null;
	WebElement login = null;
	WebElement tonalButton = null;
	WebElement okButton = null;
	WebElement searchButton = null;
	WebElement outputExcel = null;
	WebElement exportAll = null;
	WebElement export = null;
	ExcelReader excel = null;
	private boolean elementIsClickable= false;
	String projectPath = System.getProperty("user.dir");
	
	public RunPageLogin() {
		setUp();
		createLogin();
		renameFileAndCopyFileAndReadExcel();
		TearDown();
	}
	public void setUp(){
		//H:\Share Folder\Tonal\File SetUp\drivers
			System.setProperty("webdriver.chrome.driver","H:/Share Folder/Tonal/File SetUp/drivers/chromedriver.exe");

		
		ChromeOptions options = new ChromeOptions();
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", projectPath+"\\ExcelFile");
		options.setExperimentalOption("prefs", chromePrefs);
		options.addArguments("--headless", "--disable-gpu", "--window-size=1920,1200","--ignore-certificate-errors");
		DesiredCapabilities cap = DesiredCapabilities.chrome();
		cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
		cap.setCapability(ChromeOptions.CAPABILITY, options);
		driver = new ChromeDriver(cap);
		driver.get("http://10.161.168.22");
		//driver.close();
	}

	public void createLogin() {
		WebElement uname = driver.findElement(By.name("uname")); 
		WebElement psw = driver.findElement(By.name("psw"));
		uname.sendKeys("vietn");
		psw.sendKeys("biz285");
		//.xpath("//img[contains(@src,'images/out.gif')]")
		WebElement buttonLogin = driver.findElement(By.xpath("//img[contains(@src,'images/out.gif')]"));
		buttonLogin.click();

		//Go to Stations Option
		String goTo_Station = "http://10.161.168.22/index.php?_action=inventory&_operate=stations";

		//Goto Ant's Warehouse
		String WareHouse = "http://10.161.168.22/index.php?_action=inventory&_operate=warehouse";
		//driver.get(goTo_Station);
		driver.get(WareHouse);
		WebElement manufacture = driver.findElement(By.name("manufacturer_id_text"));
		manufacture.sendKeys("TONAL");
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("multi_select.multi_select('_multi_select_manufacturer_f1', 'manufacturer_id', '_multi_select_manufacturer', 'manufacturer');");
		WebDriverWait wait = new WebDriverWait(driver,20);
		WebElement tonalCheck = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@id='3' and @value='3']")));
		tonalCheck.click();
		js.executeScript("multi_select.checkbox_close('_multi_select_manufacturer_f1', '_multi_select_manufacturer', 'manufacturer_id', 'Manufacturer')");
		js.executeScript("queryrst();");
		String Station_Tab = driver.getWindowHandle();
		//WebDriver temp = driver;
		WebElement output_excel = driver.findElement(By.xpath("//input[@id='excel' and @name='excel']"));
		Point pt = output_excel.getLocation();
		int NumberX=pt.getX();
		int NumberY=pt.getY();
		Actions act= new Actions(driver);
		act.moveByOffset(NumberX+1, NumberY).click().build().perform();

		driver.switchTo().activeElement();

		WebElement outputAll = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@id='exportExcelRadio' and @value='all']")));
		outputAll.click();

		WebElement export = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'ui-button-text') and contains(text(), 'Export')]")));
		export.click();
		try {
			Thread.sleep(5000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void renameFileAndCopyFileAndReadExcel() {
		//final File folder = new File("H:/Share Folder/Tonal/Inventory History/ExcelFile");
		File[] listOfFiles = new File(projectPath+"\\ExcelFile").listFiles();
		//while(listOfFiles.length == 0);
		
		File dest = new File("H:/Share Folder/Tonal/Inventory History/AutoDownLoadRecord");
		try {
		    FileUtils.copyToDirectory(listOfFiles[0], dest);
		} catch (IOException e) {
		    e.printStackTrace();
		}
		File newFile = new File(projectPath+"/ExcelFile/input.xls");
		listOfFiles[0].renameTo(newFile);
		LocalDate date = LocalDate.now();
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
		try {
			excel = new ExcelReader(new File("H:/Share Folder/Tonal/Inventory History/AutoFileTotalRecord/"+date.format(formatter)+".xlsx"));
		} catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			Thread.sleep(3000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	public void TearDown() {
		driver.quit();
	}
	
	
}
