package com.selenium.base;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;

import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class BaseClass_WebElement_Methods {

	public static WebDriver driver;
	public static Alert sAlert;
	public static Actions objActions;
	public static Robot robot;

	public static Select sSelect;

	public static JavascriptExecutor executor;

	public static void methodToClick(WebElement webelement) {

		webelement.click();

	}

	public static void methodToSendkeys(WebElement webelement, String key) {

		webelement.sendKeys(key);
	}

	public static WebDriver browserLaunch(String browser) {

		if (browser.equalsIgnoreCase("CHROME")) {

			System.setProperty("webdriver.chrome.driver", "F:\\Eclipse Workarea\\BaseCLass\\Driver\\chromedriver.exe");

			driver = new ChromeDriver();
		} else if (browser.equalsIgnoreCase("FIREFOX")) {

			System.setProperty("webdriver.gecko.driver", "F:\\Eclipse Workarea\\BaseCLass\\Driver\\chromedriver.exe");

			WebDriver driver = new FirefoxDriver();
		} else {
			System.out.println("Invalid browser name");
		}
		driver.manage().window().maximize();
		return driver;
	}

	public static void browserClose() {

		driver.close();

	}

	public static void browserQuit() {
		driver.quit();
	}

	public static void openURL(String url) {

		driver.get(url);

	}

	public static void takeScreenshots(WebDriver driver, String destinationPath) throws IOException {

		TakesScreenshot Ts = (TakesScreenshot) driver; // narrow casting
		File source = Ts.getScreenshotAs(OutputType.FILE);

		File destination = new File(destinationPath);

		FileUtils.copyFile(source, destination);

	}

	public static void explicitElementToBeClickable(WebDriver driver, WebElement webelement, int seconds) {

		WebDriverWait waitQty = new WebDriverWait(driver, seconds);
		waitQty.until(ExpectedConditions.elementToBeClickable(webelement));
	}

	public static void methodToimplicitlyWait(WebDriver driver, int seconds) {

		driver.manage().timeouts().implicitlyWait(seconds, TimeUnit.SECONDS);

	}

	public static void explicitVisibilityOf(WebDriver driver, WebElement webelement, int seconds) {
		// TODO Auto-generated method stub

		WebDriverWait wait = new WebDriverWait(driver, seconds);
		wait.until(ExpectedConditions.visibilityOf(webelement));

	}

	public static void methodToNavigateTo(WebDriver driver, String url) {

		driver.navigate().to(url);

	}

	public static void methodToNavigateBack(WebDriver driver) {

		driver.navigate().back();

	}

	public static void methodToNavigateForward(WebDriver driver) {

		driver.navigate().forward();

	}

	public static void methodToRefresh(WebDriver driver) {

		driver.navigate().refresh();

	}

	public static void methodAcceptAlert(WebDriver driver) {

		sAlert = driver.switchTo().alert();
		sAlert.accept();

	}

	public static void methodDismissAlert(WebDriver driver) {

		sAlert = driver.switchTo().alert();
		sAlert.dismiss();

	}

	public static void methodActionsMoveToElement(WebDriver driver, WebElement webelement) {

		objActions = new Actions(driver);

		objActions.moveToElement(webelement).build().perform();

	}

	public static void methodActionsClick(WebDriver driver, WebElement webelement) {

		objActions = new Actions(driver);

		objActions.click(webelement).build().perform();

	}

	private void methodActionsContectClick(WebDriver driver, WebElement webelement) {

		objActions = new Actions(driver);

		objActions.contextClick(webelement).build().perform();

	}

	public static void methodActionsclickAndHold(WebDriver driver, WebElement webelement) {
		objActions = new Actions(driver);

		objActions.clickAndHold(webelement).build().perform();

	}

	public static void methodRobotVK_DOWN() throws AWTException {

		robot = new Robot();
		robot.keyPress(KeyEvent.VK_DOWN);
		robot.keyRelease(KeyEvent.VK_DOWN);

		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);

	}

	public static void methodToFrame(WebDriver driver, String framename) {
		driver.switchTo().frame(framename);
	}

	public static void methodToSelectByIndex(WebElement webelement, int index) {

		sSelect = new Select(webelement);

		sSelect.selectByIndex(index);

	}

	public static void methodToSelectByValue(WebElement webelement, String value) {

		sSelect = new Select(webelement);

		sSelect.selectByValue(value);

	}

	public static boolean methodToIsEnabled(WebElement webelement) {

		boolean enabled = webelement.isEnabled();
		return enabled;

	}

	public static boolean methodToIsDiaplayed(WebElement webelement) {

		boolean displayed = webelement.isDisplayed();

		return displayed;

	}

	public static boolean methodToIsSelected(WebElement webelement) {

		boolean selected = webelement.isSelected();

		return selected;

	}

	public static boolean methodIsMultiple(WebElement webElement) {

		sSelect = new Select(webElement);

		boolean multiple = sSelect.isMultiple();
		return multiple;

	}

	public static List<WebElement> getOption(WebElement webelement) {

		List<WebElement> multiOptions = null;
		sSelect = new Select(webelement);
		if (methodIsMultiple(webelement)) {
			// getting all the options from the dropdown
			multiOptions = sSelect.getOptions();

		}
		return multiOptions;

	}

	public static String methodgetTitle(WebDriver driver) {

		String title = driver.getTitle();
		return title;

	}

	public static String methodgetCurrentUrl(WebDriver driver) {

		String currentUrl = driver.getCurrentUrl();
		return currentUrl;
	}

	public static String methodgetText(WebElement webelement) {

		String text = webelement.getText();
		return text;

	}

	public static String methodGetAttribute(WebElement webelement, String attributename) {

		String attribute = webelement.getAttribute(attributename);
		return attribute;

	}

	public static List<WebElement> methodGetAllSelected(WebElement webelement) {

		List<WebElement> allSelectedOptions = null;
		sSelect = new Select(webelement);
		if (methodIsMultiple(webelement)) {
			allSelectedOptions = sSelect.getAllSelectedOptions();
		}
		return allSelectedOptions;
	}

	public static WebElement methodgetFirstSelectedOption(WebElement webelement) {

		sSelect = new Select(webelement);
		WebElement firstSelectedOption = sSelect.getFirstSelectedOption();
		return firstSelectedOption;

	}

	public static void methodScrollDownbyPosition(WebDriver driver, int point) {

		executor = (JavascriptExecutor) driver;

		executor.executeScript("window.scroll(0,point)", "");

	}

	public static void methodScrollUpPosition(WebDriver driver, int point) {
		executor = (JavascriptExecutor) driver;

		executor.executeScript("window.scroll(0,-point)", "");

	}

	public static void methodToScrollBottomOfPage(WebDriver driver) {
		executor = (JavascriptExecutor) driver;

		// To go to the bottom of the page.

		executor.executeScript("window.scrollTo(0,document.body.scrollHeight)", "");

	}

	public static void methodToScrollTop(WebDriver driver) {

		executor = (JavascriptExecutor) driver;
		executor.executeScript("window.scroll(0,0)", "");
	}

	public static void methodToScrollWebElement(WebDriver driver, WebElement webelement) {

		executor = (JavascriptExecutor) driver;
		executor.executeScript("arguments[0].scrollIntoView(true);", webelement);

	}
	
	public static String methodGetWindowHandle(WebDriver driver) {
	
		String windowHandle = driver.getWindowHandle();
		return windowHandle;

	}
	
	public static Set<String> methodGetwindowHandles(WebDriver driver) {
		
		Set<String> windowHandles = driver.getWindowHandles();
		return windowHandles;

	}
	

}
