package com.selenium.pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class Order_Confirm_POM {
	
	
	public Order_Confirm_POM(WebDriver driver2) {
		
		this.driver=driver2;
		PageFactory.initElements(driver2, this);	
		
	}
	public static WebDriver driver;
	

	

	@FindBy(xpath="(//a[@class='sf-with-ul'])[1]")
	private WebElement women;
	
	public static WebDriver getDriver() {
		return driver;
	}

	public WebElement getWomen() {
		return women;
	}

	public WebElement getImgClick() {
		return imgClick;
	}

	public WebElement getQty() {
		return Qty;
	}

	public WebElement getSize() {
		return size;
	}

	public WebElement getColor() {
		return color;
	}

	public WebElement getAddtocart() {
		return addtocart;
	}

	public WebElement getProToCheckout() {
		return proToCheckout;
	}

	public WebElement getProToCheckout2() {
		return proToCheckout2;
	}

	public WebElement getName() {
		return name;
	}

	public WebElement getPassword() {
		return password;
	}

	public WebElement getSignin() {
		return signin;
	}

	public WebElement getCheckout() {
		return checkout;
	}

	public WebElement getiAgree() {
		return iAgree;
	}

	public WebElement getProcesscarrier() {
		return processcarrier;
	}

	public WebElement getPaybycheck() {
		return paybycheck;
	}

	public WebElement getConfirmationorder() {
		return confirmationorder;
	}

	@FindBy(xpath="(//a[@itemprop='url'])[2]")
	private WebElement imgClick;
	
	//replace-2x img-responsive
	
	@FindBy(xpath="//a[@class='btn btn-default button-plus product_quantity_up']")
	private WebElement Qty;
	
	@FindBy(id="group_1")
	private WebElement size;
	
	@FindBy(xpath="//a[@title='Orange']")
	private WebElement color;
	
	@FindBy(xpath="//button[@class='exclusive']")
	private WebElement addtocart;
	
	@FindBy(xpath="//a[@title='Proceed to checkout']")
	private WebElement proToCheckout;
	
	@FindBy(xpath="(//a[@title='Proceed to checkout'])[2]")
	private WebElement proToCheckout2;
	
	
	@FindBy(name="email")
	private WebElement name;
	
	@FindBy(xpath="//input[@data-validate='isPasswd']")
	private WebElement password;
	
	@FindBy(xpath="(//button[@type='submit'])[3]")	
	private WebElement signin;
	
	@FindBy(xpath="(//button[@type='submit'])[2]")
	private WebElement checkout;
	
	@FindBy(xpath="(//input[@type='checkbox'])")
	private WebElement iAgree;
	
	@FindBy(xpath="(//button[@type='submit'])[2]")
	private WebElement processcarrier;
	
	
	@FindBy(xpath="//a[@title='Pay by check.']")
	private WebElement paybycheck;
	
	@FindBy(xpath="(//button[@type='submit'])[2]")
	private WebElement confirmationorder;
	
	
	
	
	
	
	
	
	
	
	
	
	
			
			
	

}
