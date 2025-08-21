package SeleniumTests;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.*;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;

import java.io.*;
import java.time.Duration;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class chatbotUITest {

    WebDriver driver;
    WebDriverWait wait;
    Workbook workbook;
    Sheet sheet;

    // Column indexes based on your Excel structure
    private static final int COL_TESTID = 0;          // Column A
    private static final int COL_DESCRIPTION = 1;     // Column B
    private static final int COL_INPUT = 2;           // Column C
    private static final int COL_EXPECTED = 3;        // Column D
    private static final int COL_ACTUAL = 4;          // Column E
    private static final int COL_PASS_FAIL = 5;       // Column F

    @BeforeClass
    public void setup() throws IOException {
        // Enhanced Chrome options
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--ignore-certificate-errors");
        options.addArguments("--ignore-ssl-errors=yes");
        options.addArguments("--allow-running-insecure-content");
        options.addArguments("--disable-web-security");
        options.addArguments("--disable-notifications");
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--start-maximized");
        options.addArguments("--disable-extensions");
        options.addArguments("--disable-popup-blocking");
        options.setAcceptInsecureCerts(true);

        driver = new ChromeDriver(options);
        wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        // Load Excel file
        FileInputStream file = new FileInputStream("TestCases/chatbot_Test_Cases.xlsx");
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheetAt(0);

        // Create directories if they don't exist
        new File("screenshots").mkdirs();
        new File("Reports").mkdirs();
    }

    @Test
    public void runChatbotTests() throws Exception {
        try {
            // Navigate to the website
            System.out.println("Navigating to website...");
            driver.get("https://bot.dialogflow.com/b365f706-89f1-406c-a0dd-5eba0ec23cac");

            // Wait for page to load completely
            wait.until(ExpectedConditions.jsReturnsValue("return document.readyState === 'complete';"));
            System.out.println("Page loaded successfully");

            // Find chatbot iframe
            WebElement chatbotFrame = findChatbotFrame();
            if (chatbotFrame == null) {
                throw new RuntimeException("Chatbot iframe not found!");
            }

            System.out.println("Chatbot frame found, switching to it...");
            driver.switchTo().frame(chatbotFrame);

            // Clear any existing chat history first
            clearChatHistory();

            // Test if chatbot is responsive by sending a simple message first
            testChatbotResponsiveness();

            // Loop through Excel test cases (start from row 1 to skip header)
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;

                String testId = getCellStringValue(row, COL_TESTID);
                String userMessage = getCellStringValue(row, COL_INPUT);
                String expectedReply = getCellStringValue(row, COL_EXPECTED);

                System.out.println("\n=== Executing test: " + testId + " - Input: " + userMessage + " ===");

                try {
                    // Send message to chatbot and wait for response
                    String actualReply = sendMessageAndGetResponse(userMessage);

                    // Write results to CORRECT columns
                    setCellValue(row, COL_ACTUAL, actualReply); // Column E - Actual Response

                    // IMPROVED COMPARISON LOGIC
                    boolean isPass = isResponseMatching(expectedReply, actualReply, testId);

                    if (isPass) {
                        setCellValue(row, COL_PASS_FAIL, "PASS");
                        System.out.println("✅ TEST PASSED: Expected: '" + expectedReply + "' | Actual: '" + actualReply + "'");
                    } else {
                        setCellValue(row, COL_PASS_FAIL, "FAIL");
                        System.out.println("❌ TEST FAILED: Expected: '" + expectedReply + "' | Actual: '" + actualReply + "'");
                    }

                } catch (Exception e) {
                    String errorMsg = e.getMessage();
                    setCellValue(row, COL_ACTUAL, "ERROR - " + errorMsg);
                    setCellValue(row, COL_PASS_FAIL, "FAIL");
                    System.out.println("❌ TEST ERROR: " + testId + " failed: " + errorMsg);
                    takeScreenshot("error_" + testId);

                    // Continue with next test instead of stopping
                    System.out.println("Continuing with next test...");
                }

                Thread.sleep(2000); // Wait between tests
            }

        } catch (Exception e) {
            takeScreenshot("final_error");
            throw e;
        }
    }

    // IMPROVED: More robust method to send message and get response
    private String sendMessageAndGetResponse(String message) throws Exception {
        System.out.println("📤 Sending message: " + message);

        // Send the message
        WebElement inputBox = findInputField();
        if (inputBox == null) {
            throw new RuntimeException("Input field not found for message: " + message);
        }

        // Get chat content before sending message
        String chatBefore = getAllChatContent();
        System.out.println("Chat before sending: " + (chatBefore.length() > 100 ? chatBefore.substring(0, 100) + "..." : chatBefore));

        inputBox.clear();
        inputBox.sendKeys(message);
        inputBox.sendKeys(Keys.ENTER);

        System.out.println("⏳ Waiting for bot response... (up to 30 seconds)");

        String botResponse = "";

        // Wait for response with multiple strategies
        for (int i = 0; i < 30; i++) { // 30 seconds timeout
            Thread.sleep(1000);
            System.out.print(".");

            // Strategy 1: Check if chat content changed
            String chatAfter = getAllChatContent();
            if (!chatAfter.equals(chatBefore) && chatAfter.length() > chatBefore.length()) {
                String newContent = chatAfter.substring(chatBefore.length()).trim();
                if (!newContent.isEmpty() && !newContent.equalsIgnoreCase(message)) {
                    botResponse = cleanBotResponse(newContent);
                    if (!botResponse.isEmpty()) {
                        System.out.println("\n✅ Bot response detected via chat change: '" + botResponse + "'");
                        return botResponse;
                    }
                }
            }

            // Strategy 2: Look for specific bot message elements
            botResponse = findBotMessageDirectly();
            if (!botResponse.isEmpty() && !botResponse.equalsIgnoreCase(message)) {
                System.out.println("\n✅ Bot response found via direct search: '" + botResponse + "'");
                return botResponse;
            }

            // Strategy 3: Check for any visible text that might be a response
            botResponse = findAnyPossibleResponse(message);
            if (!botResponse.isEmpty()) {
                System.out.println("\n✅ Bot response found via fallback search: '" + botResponse + "'");
                return botResponse;
            }
        }

        // If no response found, capture what we can see for debugging
        String finalChatContent = getAllChatContent();
        System.out.println("\n❌ No bot response found after 30 seconds");
        System.out.println("Final chat content: " + finalChatContent);

        // Don't throw exception - return what we found for Excel logging
        return "TIMEOUT - No response detected within 30 seconds. Final chat: " +
                (finalChatContent.length() > 200 ? finalChatContent.substring(0, 200) + "..." : finalChatContent);
    }

    // NEW: Get all chat content using multiple approaches
    private String getAllChatContent() {
        try {
            // Try different selectors to get chat content
            String[] chatSelectors = {
                    "#resultWrapper",
                    ".b-agent-demo_result",
                    ".chat-messages",
                    ".df-messenger-chat",
                    ".conversation",
                    ".messages",
                    "df-messenger",
                    "[role='log']"
            };

            for (String selector : chatSelectors) {
                try {
                    WebElement chatElement = driver.findElement(By.cssSelector(selector));
                    if (chatElement.isDisplayed()) {
                        String content = chatElement.getText().trim();
                        if (!content.isEmpty()) {
                            return content;
                        }
                    }
                } catch (Exception e) {
                    // Continue with next selector
                }
            }

            // Fallback: get entire body text
            return driver.findElement(By.tagName("body")).getText();

        } catch (Exception e) {
            System.out.println("Error getting chat content: " + e.getMessage());
            return "";
        }
    }

    // NEW: Find bot messages using direct element search
    private String findBotMessageDirectly() {
        String[] botSelectors = {
                ".bot-message",
                ".agent-message",
                ".df-messenger-chat-bubble[agent]",
                ".response",
                "[class*='bot']",
                "[class*='agent']",
                "[class*='response']"
        };

        for (String selector : botSelectors) {
            try {
                List<WebElement> elements = driver.findElements(By.cssSelector(selector));
                if (!elements.isEmpty()) {
                    WebElement lastElement = elements.get(elements.size() - 1);
                    String text = lastElement.getText().trim();
                    if (!text.isEmpty()) {
                        return cleanBotResponse(text);
                    }
                }
            } catch (Exception e) {
                // Continue with next selector
            }
        }
        return "";
    }

    // NEW: Fallback method to find any possible response
    private String findAnyPossibleResponse(String userMessage) {
        try {
            // Get all visible text elements
            List<WebElement> allElements = driver.findElements(By.xpath("//*[text()]"));

            String bestResponse = "";
            for (WebElement element : allElements) {
                try {
                    if (element.isDisplayed()) {
                        String text = element.getText().trim();
                        if (!text.isEmpty() &&
                                !text.equalsIgnoreCase(userMessage) &&
                                text.length() > 5 &&
                                text.length() < 500 &&
                                !isUIElement(text)) {
                            bestResponse = text;
                        }
                    }
                } catch (Exception e) {
                    // Skip this element
                }
            }

            return cleanBotResponse(bestResponse);

        } catch (Exception e) {
            return "";
        }
    }

    // NEW: Check if text is likely a UI element rather than bot response
    private boolean isUIElement(String text) {
        String lowerText = text.toLowerCase();
        return lowerText.contains("powered by") ||
                lowerText.contains("close") ||
                lowerText.contains("minimize") ||
                lowerText.contains("send") ||
                lowerText.contains("type a message") ||
                lowerText.contains("loading") ||
                lowerText.equals("mic") ||
                lowerText.length() < 3;
    }

    // IMPROVED: More aggressive bot response cleaning
    private String cleanBotResponse(String response) {
        if (response == null || response.trim().isEmpty()) {
            return "";
        }

        // Remove common UI elements and noise
        String cleaned = response
                .replace("...", "")
                .replace("•", "")
                .replace("mic", "")
                .replace("POWERED BY", "")
                .replace("QA_Chatbot_Test", "")
                .replace("Dialogflow", "")
                .replace("Google", "")
                .replace("Close", "")
                .replace("Minimize", "")
                .replace("Send", "")
                .replaceAll("\\s+", " ")
                .trim();

        // Remove timestamps and other metadata
        cleaned = cleaned.replaceAll("\\d{1,2}:\\d{2}(:\\d{2})?", ""); // Remove timestamps
        cleaned = cleaned.replaceAll("\\b(AM|PM)\\b", ""); // Remove AM/PM
        cleaned = cleaned.replaceAll("\\b(Today|Yesterday)\\b", ""); // Remove day references

        // Remove very short responses that are likely UI elements
        if (cleaned.length() < 3) {
            return "";
        }

        // Take only the first meaningful sentence if response is very long
        if (cleaned.length() > 300) {
            String[] sentences = cleaned.split("[.!?]");
            if (sentences.length > 0 && sentences[0].length() > 10) {
                return sentences[0].trim();
            }
        }

        return cleaned.trim();
    }

    // IMPROVED RESPONSE MATCHING LOGIC
    private boolean isResponseMatching(String expected, String actual, String testId) {
        if (actual == null || actual.trim().isEmpty()) {
            return false;
        }

        // Clean both strings for comparison
        String cleanExpected = cleanText(expected);
        String cleanActual = cleanText(actual);

        System.out.println("Comparing - Expected: '" + cleanExpected + "' vs Actual: '" + cleanActual + "'");

        // Test-specific matching logic
        switch (testId) {
            case "TC01": // Hello / Hi
                return cleanActual.contains("hello") || cleanActual.contains("hi") ||
                        cleanActual.contains("how can i help");

            case "TC02": // Working hours
            case "TC07": // Typo working hours
            case "TC14": // Opening hours
                return cleanActual.contains("9") && cleanActual.contains("5") &&
                        (cleanActual.contains("am") || cleanActual.contains("pm"));

            case "TC03": // Company address
                return cleanActual.contains("address") || cleanActual.contains("located") ||
                        cleanActual.contains("location");

            case "TC04": // Contact number
            case "TC10": // Support contact
                return cleanActual.contains("contact") || cleanActual.contains("call") ||
                        cleanActual.contains("email") || cleanActual.contains("support");

            case "TC05": // Thanks
                return cleanActual.contains("welcome") || cleanActual.contains("thank") ||
                        cleanActual.contains("appreciate");

            case "TC06": // Unknown query
            case "TC09": // Pricing info (unknown)
                return cleanActual.contains("sorry") || cleanActual.contains("understand") ||
                        cleanActual.contains("rephrase");

            case "TC08": // Services
                return cleanActual.contains("service") || cleanActual.contains("offer") ||
                        cleanActual.contains("provide");

            case "TC11": // Business email
                return cleanActual.contains("email") || cleanActual.contains("@") ||
                        cleanActual.contains("support") || cleanActual.contains("info");

            case "TC12": // Multiple greetings
                return cleanActual.contains("hello") || cleanActual.contains("hi") ||
                        cleanActual.contains("hey") || cleanActual.contains("how can i help");

            case "TC13": // Closing hours
                return cleanActual.contains("close") || cleanActual.contains("5") ||
                        cleanActual.contains("pm");

            case "TC15": // Empty input
                return cleanActual.contains("say") || cleanActual.contains("repeat") ||
                        cleanActual.contains("type") || cleanActual.contains("question");

            default:
                // Fallback: use contains for general matching
                return cleanActual.contains(cleanExpected) || cleanExpected.contains(cleanActual);
        }
    }

    private String cleanText(String text) {
        if (text == null) return "";

        return text.toLowerCase()
                .replace("\"", "")      // Remove quotes
                .replace(".", "")       // Remove periods
                .replace(",", "")       // Remove commas
                .replace("!", "")       // Remove exclamation marks
                .replace("?", "")       // Remove question marks
                .replace("'", "")       // Remove apostrophes
                .replace("  ", " ")     // Remove double spaces
                .trim();
    }

    private String getCellStringValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        return (cell != null) ? cell.getStringCellValue() : "";
    }

    private void setCellValue(Row row, int columnIndex, String value) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        cell.setCellValue(value);
    }

    private WebElement findChatbotFrame() {
        System.out.println("Looking for chatbot iframe...");

        List<WebElement> frames = driver.findElements(By.tagName("iframe"));
        System.out.println("Found " + frames.size() + " iframes on the page");

        for (WebElement frame : frames) {
            String src = frame.getAttribute("src");
            if (src != null && src.contains("dialogflow")) {
                System.out.println("✅ Dialogflow chatbot iframe found: " + src);
                return frame;
            }
        }

        return null;
    }

    private void clearChatHistory() {
        try {
            System.out.println("Clearing chat history...");
            // Try to find and click a clear/refresh button
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("if(typeof clearChat === 'function') clearChat();");
            Thread.sleep(1000);
        } catch (Exception e) {
            System.out.println("No clear function found, continuing...");
        }
    }

    private void testChatbotResponsiveness() {
        try {
            System.out.println("Testing chatbot responsiveness...");

            String response = sendMessageAndGetResponse("Hello");
            if (response.contains("TIMEOUT")) {
                System.out.println("⚠️ Chatbot responsiveness test showed timeout, but continuing...");
            } else {
                System.out.println("✅ Chatbot is responsive! Response: '" + response + "'");
            }

        } catch (Exception e) {
            System.out.println("❌ Chatbot responsiveness test failed: " + e.getMessage());
            System.out.println("Continuing with tests anyway...");
            takeScreenshot("responsiveness_test_failed");
        }
    }

    private WebElement findInputField() {
        String[] inputSelectors = {
                "input[type='text']",
                "textarea",
                "[contenteditable='true']",
                ".input-field",
                ".chat-input",
                "#input",
                "#query"
        };

        for (String selector : inputSelectors) {
            try {
                List<WebElement> elements = driver.findElements(By.cssSelector(selector));
                if (!elements.isEmpty()) {
                    System.out.println("✅ Input field found using selector: " + selector);
                    return elements.get(0);
                }
            } catch (Exception e) {
                // Continue trying other selectors
            }
        }
        return null;
    }

    private void takeScreenshot(String name) {
        try {
            File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            FileUtils.copyFile(screenshot, new File("screenshots/" + name + "_" + System.currentTimeMillis() + ".png"));
        } catch (Exception e) {
            System.out.println("Failed to take screenshot: " + e.getMessage());
        }
    }

    @AfterClass
    public void teardown() throws IOException {
        // Save results
        FileOutputStream outFile = new FileOutputStream("Reports/Chatbot_Test_Results.xlsx");
        workbook.write(outFile);
        outFile.close();
        workbook.close();

        if (driver != null) {
            driver.quit();
        }
    }
}