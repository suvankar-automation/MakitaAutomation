package utility;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

public class ScreenShot 
{
	public static String captureScreenshot(WebDriver driver,String screenshotName)
	{
		try
		{
			//TakesScreenshot is an interface. we are typecasting WebDriver reference to TakesScreenshot
			TakesScreenshot ts=(TakesScreenshot)driver; 
			File source=ts.getScreenshotAs(OutputType.FILE);
			
			String screenshotLocation = System.getProperty("user.dir")+"/ScreenShots/";
			Calendar calendar = Calendar.getInstance();
			SimpleDateFormat formater = new SimpleDateFormat("dd_MM_yyyy_hh_mm_ss");
			String FullScreenshotName = screenshotLocation+screenshotName+"_"+formater.format(calendar.getTime())+".png";
			
			File destFile = new File(FullScreenshotName);
			FileUtils.copyFile(source, destFile);
			
			return FullScreenshotName;
		}
		catch(Exception e)
		{
			System.out.println("Exception while taking screenshot:"+e.getMessage());
			return e.getMessage();
		}
	}
}
