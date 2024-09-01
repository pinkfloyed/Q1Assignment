import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Optional;

public class Main {

    public static void main(String[] args) {
        WebDriver driver = null;
        FileInputStream file = null;
        FileOutputStream outFile = null;
        Workbook workbook = null;
        try {
            // Set up the WebDriver (replace with your path to chromedriver)
            System.setProperty("webdriver.chrome.driver", "C:\\Users\\pinki99\\Downloads\\chromedriver-win64\\chromedriver.exe");
            driver = new ChromeDriver();

            // Open the Excel file
            file = new FileInputStream(new File("C:\\Users\\pinki99\\Downloads\\4BeatsQ1.xlsx"));
            workbook = new XSSFWorkbook(file);

            // Get today's day of the week
            String today = new SimpleDateFormat("EEEE").format(new Date());
            Sheet sheet = workbook.getSheet(today); // Access the sheet corresponding to today's day

            // Loop through the rows and process keywords
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Start from row 1 (assuming row 0 is the header)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell keywordCell = row.getCell(1); // Assuming keywords are in the second column
                    if (keywordCell != null) {
                        String keyword = keywordCell.getStringCellValue();

                        // Search the keyword on Google
                        driver.get("https://www.google.com");
                        WebElement searchBox = driver.findElement(By.name("q"));
                        searchBox.sendKeys(keyword);
                        searchBox.submit();

                        // Extract search results text
                        List<WebElement> searchResults = driver.findElements(By.cssSelector("h3"));
                        Optional<String> longest = searchResults.stream()
                                .map(WebElement::getText)
                                .filter(text -> !text.isEmpty())
                                .max((a, b) -> Integer.compare(a.length(), b.length()));
                        Optional<String> shortest = searchResults.stream()
                                .map(WebElement::getText)
                                .filter(text -> !text.isEmpty())
                                .min((a, b) -> Integer.compare(a.length(), b.length()));

                        // Write longest and shortest to Excel
                        Cell longestCell = row.createCell(2); // Column C for longest option
                        Cell shortestCell = row.createCell(3); // Column D for shortest option
                        longestCell.setCellValue(longest.orElse("N/A"));
                        shortestCell.setCellValue(shortest.orElse("N/A"));
                    }
                }
            }

            // Save the updated Excel file
            file.close(); // Close the input stream before writing the output
            outFile = new FileOutputStream(new File("C:\\Users\\pinki99\\Downloads\\4BeatsQ1_updated.xlsx")); // Ensure a valid path
            workbook.write(outFile);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                // Clean up resources
                if (driver != null) {
                    driver.quit();
                }
                if (workbook != null) {
                    workbook.close();
                }
                if (file != null) {
                    file.close();
                }
                if (outFile != null) {
                    outFile.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
