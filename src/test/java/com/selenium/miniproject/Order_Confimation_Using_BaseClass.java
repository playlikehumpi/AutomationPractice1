package com.selenium.miniproject;

import java.io.IOException;

import org.openqa.selenium.WebDriver;

import com.selenium.pom.Order_Confirm_POM;


public class Order_Confimation_Using_BaseClass extends com.selenium.base.BaseClass_WebElement_Methods  {

	public static WebDriver driver;

	public static void main(String[] args) throws InterruptedException, IOException {

		driver = browserLaunch("chrome");

		openURL("http://automationpractice.com/index.php");

		Thread.sleep(7000);

		Order_Confirm_POM order = new Order_Confirm_POM(driver);

		methodToClick(order.getWomen());

		Thread.sleep(10000);

		explicitVisibilityOf(driver, order.getImgClick(), 50);

		methodToScrollWebElement(driver, order.getImgClick());

		Thread.sleep(10000);
		explicitVisibilityOf(driver, order.getImgClick(), 50);
		methodToClick(order.getImgClick());

		Thread.sleep(7000);

		explicitElementToBeClickable(driver, order.getQty(), 30);

		methodToimplicitlyWait(driver, 4000);

		methodToClick(order.getQty());

		Thread.sleep(7000);

		methodToSelectByIndex(order.getSize(), 2);

		methodToClick(order.getColor());

		methodToClick(order.getAddtocart());

		methodToClick(order.getProToCheckout());

		methodToClick(order.getProToCheckout2());

		methodToSendkeys(order.getName(), "seltest340@gmail.com");

		methodToSendkeys(order.getPassword(), "UISDJjk");

		methodToClick(order.getSignin());

		methodToClick(order.getCheckout());

		methodToClick(order.getiAgree());

		methodToClick(order.getProcesscarrier());

		methodToClick(order.getPaybycheck());

		Thread.sleep(3000);

		methodToClick(order.getConfirmationorder());

		// To Capture Screenshot

		takeScreenshots(driver, "F:\\Eclipse Workarea\\AutomationPractice1\\Screenshots\\orderconfirmation.png");

	}
}
