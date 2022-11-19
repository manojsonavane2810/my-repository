package maven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class excel1 {

	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\Apmosys\\Desktop\\Assessment_Link.xlsx");
		InputStream fis = new FileInputStream(f);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("Sheet1");
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://money.rediff.com/gainers");
		driver.manage().timeouts().implicitlyWait(6, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		WebElement tab = driver.findElement(By.xpath("//table[@class='dataTable']"));
		WebElement tb = tab.findElement(By.tagName("tbody"));
		WebElement th = tab.findElement(By.tagName("thead"));

		List<WebElement> row = tb.findElements(By.tagName("tr"));
		
		for (int i = 0; i < 50; i++) {
			Row r=sheet.createRow(i);
			List<WebElement> col = row.get(i).findElements(By.tagName("td"));
			for (int j = 0; j < col.size(); j++) {
				Cell cell=r.createCell(j);
			//	System.out.print(col.get(j).getText());
				System.out.println(i);
				cell.setCellValue(col.get(j).getText());
			}
			System.out.println();

		}
		fis.close();
		FileOutputStream fos=new FileOutputStream(f);
		wb.write(fos);
		wb.close();
		fos.close();
		InputStream fis1 = new FileInputStream(f);
		XSSFWorkbook wb1=new XSSFWorkbook(fis1);
		XSSFSheet sheet1=wb.getSheet("Sheet1");
		System.out.print(sheet1.getRow(5).getCell(0)+"   "+sheet1.getRow(5).getCell(3));
		wb1.close();
		fos.close();
	}
	

}
