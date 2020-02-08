package com.google.tests;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class MakitaAutomatedScript {

	public static void main(String[] args) throws InterruptedException, AWTException, IOException
	{
		int i=1;
		System.setProperty("webdriver.chrome.driver","C:\\Users\\Suvankar\\Desktop\\Selenium\\chromedriver\\chromedriver_win32\\chromedriver.exe");
		String baseUrl = "http://www.makita.in/";
		//WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(baseUrl)  ; 

		driver.findElement(By.linkText("Tools")).click();
		List <WebElement> links=driver.findElements(By.xpath("//*[@class='box-category']//ul/li/a"));
		//System.out.println(links.size());
		
		for (WebElement element : links){

			String link=element.getAttribute("href");
			
			//To open a new tab         
			Robot r = new Robot();                          
			r.keyPress(KeyEvent.VK_CONTROL); 
			r.keyPress(KeyEvent.VK_T); 
			r.keyRelease(KeyEvent.VK_CONTROL); 
			r.keyRelease(KeyEvent.VK_T);
			
			//To switch to the new tab
			ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
			driver.switchTo().window(tabs.get(1));
			
			//To navigate to new link/URL in 2nd new tab  
			driver.get(link); 
			Thread.sleep(5000);
			
			boolean flag;
			do {
			List <WebElement> products=driver.findElements(By.xpath("//div[@class='product-grid']/div//a[contains(@href, 'makita')]"));
			//System.out.println(products.size());

			for (WebElement product : products){

				String link1=product.getAttribute("href");
				//System.out.println(link1);
				r.keyPress(KeyEvent.VK_CONTROL); 
				r.keyPress(KeyEvent.VK_T); 
				r.keyRelease(KeyEvent.VK_CONTROL); 
				r.keyRelease(KeyEvent.VK_T);
				
				ArrayList<String> tabs1 = new ArrayList<String>(driver.getWindowHandles());
				driver.switchTo().window(tabs1.get(2));
				
				driver.get(link1); 
				Thread.sleep(5000);
				
				//Get Product Name
				String prodName=driver.findElement(By.xpath("//*[@class='product-info']//b")).getText();
				String prodCode=driver.findElement(By.xpath("//*[@class='breadcrumb']//a[4]")).getText();
				System.out.println(prodCode);
				
				//Download Product Image
				WebElement image =driver.findElement(By.xpath("//*[@class='product-info']//img[@id='image']"));
				String imgLink= image.getAttribute("src");
				//System.out.println(imgLink);
				//Timestamp timestamp = new Timestamp(System.currentTimeMillis());
				//long tm=timestamp.getTime();
				String destinationFile = prodCode+"_"+prodName+".jpg";
				
				BufferedImage pic =null;
				try{

					URL url =new URL(imgLink);
					// Read the URL
					pic = ImageIO.read(url);

					// for jpg
					ImageIO.write(pic, "jpg",new File("D:\\Test\\"+destinationFile));

				}catch(IOException e){
					e.printStackTrace();
				}
				
				String excelPath1="D:\\Test\\TestFile.xlsx";
				ExcelHandler excel=new ExcelHandler(excelPath1,"TestTab");	
				
				//Get the Product Description
				List<WebElement> allRows=driver.findElements(By.xpath("//*[@class='attribute']/tbody/tr"));

				for (WebElement row : allRows) { 
					List<WebElement> cells = row.findElements(By.tagName("td")); 
					excel.addNewRow(i);
					// Print the contents of each cell
					int j=1;
					excel.updateCellValue(i, j-1, prodCode+"_"+prodName);
					for (WebElement cell : cells) { 
						//System.out.println(cell.getText());
						excel.updateCellValue(i, j, cell.getText());
						//excel.commit();
						j=j+1;
					}
					i=i+1;
				}
				excel.commit();
				
				driver.close();
				driver.switchTo().window(tabs.get(1));	
			}
			
			//Pagination
			//driver.get("http://makita.in/index.php?route=product/category&path=20_59");
			try {
				driver.findElement(By.xpath("//*[@class='links']/a[text()='>']")).click();
				//Thread.sleep(5000);
				flag= true;
			} catch (NoSuchElementException e){
				//System.out.println("Handled NoSuchElementException");
				flag=false;
			}
			
			}while(flag);
				
			driver.close();
			driver.switchTo().window(tabs.get(0));
		}
		
		driver.close();
	}

}
