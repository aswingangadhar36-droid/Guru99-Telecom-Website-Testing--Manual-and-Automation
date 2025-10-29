package Final_project;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.*;

public class Guru99TelecomWebsite {
	
	    WebDriver driver;
	    String excelPath ="C:\\Users\\user\\Desktop\\CustomerData.xlsx" ; 
	    @BeforeClass
	    public void setUp() throws InterruptedException {
	        driver = new ChromeDriver();
	        driver.manage().window().maximize();
	        driver.get("https://demo.guru99.com/telecom/index.html");
	        Thread.sleep(2000);
	        System.out.println("Browser launched and Guru99 Telecom site opened.");
	    }


	    // Data Provider //
	    @SuppressWarnings("deprecation")
		@DataProvider(name = "CustomerData")
	    public Object[][] getCustomerData() throws Exception {
	        FileInputStream file = new FileInputStream("C:\\Users\\user\\Desktop\\CustomerData.xlsx"); // Excel path
	        Workbook workbook = new XSSFWorkbook(file);
	        Sheet sheet = workbook.getSheetAt(0);

	        int rowCount = sheet.getPhysicalNumberOfRows() - 1; 
	        Object[][] data = new Object[rowCount][5];

	        for (int i = 1; i <= rowCount; i++) {
	            Row row = sheet.getRow(i);

	            data[i - 1][0] = row.getCell(0).getStringCellValue(); // First Name
	            data[i - 1][1] = row.getCell(1).getStringCellValue(); // Last Name
	            data[i - 1][2] = row.getCell(2).getStringCellValue(); // Email
	            data[i - 1][3] = row.getCell(3).getStringCellValue(); // Address

	            // Phone number as string (avoid scientific notation)
	            Cell phoneCell = row.getCell(4);
	            phoneCell.setCellType(CellType.STRING);
	            data[i - 1][4] = phoneCell.getStringCellValue();
	        }
	        workbook.close();
	        return data;
	    }

	  
	   

	  
	    @Test(priority = 2, dataProvider = "CustomerData")
	    public void addCustomer(String fname, String lname, String email, String address, String phone)
	            throws InterruptedException, IOException {

	    	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	        try {
	            Alert alert = driver.switchTo().alert();
	            System.out.println("Pre-existing Alert: " + alert.getText());
	            alert.accept();
	        } catch (NoAlertPresentException e) {
	            // no alert, continue
	        }

	    	WebElement addCustomerLink = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@href='addcustomer.php']")));
	    	addCustomerLink.click();
	        driver.findElement(By.cssSelector("input[id='fname']")).sendKeys(fname);
	        driver.findElement(By.xpath("//input[@name='lname']")).sendKeys(lname);
	        driver.findElement(By.name("emailid")).sendKeys(email);
	        driver.findElement(By.xpath("//textarea[@name='addr']")).sendKeys(address);
	        driver.findElement(By.name("telephoneno")).sendKeys(phone);
	        System.out.println("Adding Customer: " + fname + " | " + lname + " | " + email + " | " + address + " | " + phone);

	        driver.findElement(By.xpath("//input[@type='submit']")).click();
	        try {
	            Alert postAlert = driver.switchTo().alert();
	            System.out.println("Post-submit Alert: " + postAlert.getText());
	            postAlert.accept();
	        } catch (NoAlertPresentException e) {
	            System.out.println("No post-submit alert");
	        }

	        // Optional: Wait for success message or next page element
	        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='table-wrapper']")));
	    
	

	        Thread.sleep(2000);
	        WebElement customerId = driver.findElement(By.xpath("//div[@class='table-wrapper']"));
	        File src = customerId.getScreenshotAs(OutputType.FILE);
	        File dest = new File("./Screenshots/" + fname + "_CustomerID.png");
	        FileUtils.copyFile(src, dest);
	        System.out.println("Screenshot saved for " + fname);
	        driver.navigate().to("https://demo.guru99.com/telecom/index.html");
	    }
	    @Test(priority = 1)
	    public void verifyTitleAndURL() {
	        String pageTitle = driver.getTitle();
	        System.out.println("Page Title: " + pageTitle);
	        String pageURL = driver.getCurrentUrl();
	        System.out.println("Current URL: " + pageURL);
	        if (pageURL.equals("https://demo.guru99.com/telecom/index.html")) {
	            System.out.println("✅ Actual URL matches expected URL.");
	        } else {
	            System.out.println("❌ Actual URL does not match expected URL.");
	        }
	    }
	    
	    @Test(priority = 0)
	    public void verifyAlert() throws InterruptedException {
	        driver.findElement(By.linkText("Pay Billing")).click();
	        driver.findElement(By.xpath("//input[@type='submit']")).click();
	        Alert simpleAlert = driver.switchTo().alert();
	        System.out.println("Alert text: " + simpleAlert.getText());
	        Thread.sleep(2000);
	        simpleAlert.accept();
	        System.out.println("Alert accepted successfully.");
	        driver.navigate().back();
	    }


	    @Test(priority = 3)
	    public void copyPasteExample() throws InterruptedException {
	        driver.findElement(By.linkText("Add Tariff Plan to Customer")).click();
	        Thread.sleep(2000);
	        WebElement element1 = driver.findElement(By.id("customer_id"));
	        element1.sendKeys("536887");
	        element1.sendKeys(Keys.CONTROL + "a");
	        element1.sendKeys(Keys.CONTROL + "c");
	        WebElement element2 = driver.findElement(By.id("customer_id"));
	        element2.clear();
	        element2.sendKeys(Keys.CONTROL + "v");
	        Thread.sleep(2000);
	        System.out.println("Copied and pasted Customer ID successfully.");
	        driver.navigate().to("https://demo.guru99.com/telecom/index.html");
	    }
	    @Test(priority = 4)
	    
	    public void addtariffPlan() throws InterruptedException {
	    	
	    	driver.findElement(By.linkText("Add Tariff Plan")).click();
	    	Thread.sleep(2000);
	    	WebElement rent = driver.findElement(By.name("rental"));
	    	rent.sendKeys("100");
	    	driver.findElement(By.name("local_minutes")).sendKeys("200");
	    	driver.findElement(By.name("inter_minutes")).sendKeys("300");
	    	driver.findElement(By.name("sms_pack")).sendKeys("50");
	    	driver.findElement(By.name("minutes_charges")).sendKeys("500");
	        driver.findElement(By.name("inter_charges")).sendKeys("1000");
	    	driver.findElement(By.name("sms_charges")).sendKeys("150");
	    	driver.findElement(By.xpath("//input[@type='submit']")).click();
	    	
	    	
	    }

	    @AfterClass
	    public void tearDown() {
	        driver.quit();
	        System.out.println("Browser closed successfully.");
	    }
	}


