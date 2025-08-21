package SeleniumTests;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.*;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.By;

import java.io.*;
import java.time.Duration;
import java.util.List;

public class chatbotUITest {

    WebDriver driver;
    WebDriverWait wait;
    Workbook workbook;
    Sheet sheet;

    @BeforeClass
    public void setup() throws IOException {
        // Setup Chrome options for better stability
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-blink-features=AutomationControlled");
        options.addArguments("--disable-extensions");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--remote-allow-origins=*");

        // Initialize WebDriver
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\lniru\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        // Initialize WebDriverWait after driver is created
        wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        // Create directories if they don't exist
        createDirectoryIfNotExists("Screenshots");
        createDirectoryIfNotExists("Reports");

        // Load Excel file
        try {
            FileInputStream file = new FileInputStream("TestCases/Chatbot_Test_Cases.xlsx");
            workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(0);
            file.close();
        } catch (IOException e) {
            System.err.println("Error loading Excel file: " + e.getMessage());
            throw e;
        }
    }

    @Test
    public void runChatbotTests() throws IOException, InterruptedException {
        try {
            // Navigate to the chatbot URL
            driver.get("https://dialogflow.cloud.google.com/#/demo");

            // Wait for the page to load completely
            Thread.sleep(3000);

            // Process each test case from Excel
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                // Skip empty rows
                if (row == null || row.getCell(1) == null) {
                    continue;
                }

                String testCaseId = getCellValueAsString(row.getCell(0));
                String input = getCellValueAsString(row.getCell(1));
                String expected = getCellValueAsString(row.getCell(2));

                System.out.println("Running Test Case: " + testCaseId + " - Input: " + input);

                try {
                    // Wait for input box and send message
                    WebElement inputBox = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//input[@placeholder='Type a message' or @placeholder='Ask me anything...' or contains(@class, 'input')]")));

                    // Clear any existing text and send new input
                    inputBox.clear();
                    inputBox.sendKeys(input);
                    inputBox.sendKeys(Keys.ENTER);

                    // Wait for response
                    Thread.sleep(3000);

                    // Try multiple selectors for response
                    WebElement response = null;
                    String[] responseSelectors = {
                            ".message-text",
                            "[class*='message']",
                            "[class*='response']",
                            "[class*='bot-message']",
                            ".df-messenger-user-message + div",
                            "df-messenger-user-message ~ df-messenger-bot-message"
                    };

                    for (String selector : responseSelectors) {
                        try {
                            List<WebElement> elements = driver.findElements(By.cssSelector(selector));
                            if (!elements.isEmpty()) {
                                response = elements.get(elements.size() - 1); // Get the last response
                                break;
                            }
                        } catch (Exception e) {
                            // Continue to next selector
                        }
                    }

                    String actual = "";
                    if (response != null) {
                        actual = response.getText().trim();
                    } else {
                        actual = "No response found";
                        System.out.println("Warning: Could not find response element for test case: " + testCaseId);
                    }

                    // Write results to Excel
                    Cell actualCell = row.createCell(3);
                    actualCell.setCellValue(actual);

                    String result = actual.toLowerCase().contains(expected.toLowerCase()) ? "Pass" : "Fail";
                    Cell resultCell = row.createCell(4);
                    resultCell.setCellValue(result);

                    System.out.println("Expected: " + expected + " | Actual: " + actual + " | Result: " + result);

                    // Take screenshot if test fails
                    if (result.equals("Fail")) {
                        takeScreenshot(testCaseId);
                    }

                    // Small delay between tests
                    Thread.sleep(1000);

                } catch (Exception e) {
                    System.err.println("Error in test case " + testCaseId + ": " + e.getMessage());

                    // Write error to Excel
                    Cell actualCell = row.createCell(3);
                    actualCell.setCellValue("Error: " + e.getMessage());

                    Cell resultCell = row.createCell(4);
                    resultCell.setCellValue("Error");

                    // Take screenshot for error
                    takeScreenshot(testCaseId + "_error");
                }
            }

            // Save results to Excel
            saveResultsToExcel();

        } catch (Exception e) {
            System.err.println("Error in runChatbotTests: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    private void takeScreenshot(String testCaseId) {
        try {
            File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            File targetFile = new File("Screenshots/" + testCaseId + ".png");

            // Use proper file copy method
            try (FileInputStream fis = new FileInputStream(screenshot);
                 FileOutputStream fos = new FileOutputStream(targetFile)) {
                byte[] buffer = new byte[1024];
                int length;
                while ((length = fis.read(buffer)) > 0) {
                    fos.write(buffer, 0, length);
                }
            }

            System.out.println("Screenshot saved: " + targetFile.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("Error taking screenshot: " + e.getMessage());
        }
    }

    private void saveResultsToExcel() {
        try {
            FileOutputStream outFile = new FileOutputStream("Reports/Chatbot_Test_Results.xlsx");
            workbook.write(outFile);
            outFile.close();
            System.out.println("Test results saved to Reports/Chatbot_Test_Results.xlsx");
        } catch (IOException e) {
            System.err.println("Error saving results to Excel: " + e.getMessage());
        }
    }

    private void createDirectoryIfNotExists(String dirPath) {
        File directory = new File(dirPath);
        if (!directory.exists()) {
            directory.mkdirs();
            System.out.println("Created directory: " + dirPath);
        }
    }

    @AfterClass
    public void teardown() {
        try {
            if (workbook != null) {
                workbook.close();
            }
        } catch (IOException e) {
            System.err.println("Error closing workbook: " + e.getMessage());
        }

        if (driver != null) {
            driver.quit();
        }

        System.out.println("Test execution completed and resources cleaned up.");
    }
}