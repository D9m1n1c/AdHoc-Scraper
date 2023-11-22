const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const targetURL = "https://lihkg.com/thread/3466186/page/1"; // You can change this URL
const totalPages = 10; // The total number of pages you want to scrape
const threadID = targetURL.match(/thread\/(\d+)/)[1];

(async () => {
  // Launch a non-headless browser
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: false,
  });

  const allComments = [];

  // Loop through pages
  for (let pageIdx = 1; pageIdx <= totalPages; pageIdx++) {
    const page = await browser.newPage();

    // Navigate to the target page
    const pageURL = targetURL.replace(/page\/\d+/, `page/${pageIdx}`);
    await page.goto(pageURL, {
      waitForSelector: "#rightPanel div._4SkbmyWbShmINbRPY-Jps",
    });
    await page.waitForTimeout(5000);

    // Function to extract comments from a specific element index
    const extractComments = async (index) => {
      const usernameSelector = `#rightPanel div._4SkbmyWbShmINbRPY-Jps div:nth-child(${index}) div small span:nth-child(2)`;
      const dateSelector = `#rightPanel div._4SkbmyWbShmINbRPY-Jps div:nth-child(${index}) div small span:nth-child(4)`;
      const contentSelector = `#rightPanel div._4SkbmyWbShmINbRPY-Jps div:nth-child(${index}) div._2cNsJna0_hV8tdMj3X6_gJ`;
      const username = await page.evaluate((selector) => {
        const element = document.querySelector(selector);
        return element ? element.textContent.trim() : "";
      }, usernameSelector);

      const date = await page.evaluate((selector) => {
        const element = document.querySelector(selector);
        return element ? element.title : "";
      }, dateSelector);

      const content = await page.evaluate((selector) => {
        const element = document.querySelector(selector);
        return element ? element.textContent.trim() : "";
      }, contentSelector);

      return { username, date, content };
    };

    // Extract comments for the current page

    for (let i = 2; i <= 50; i++) {
      try {
        const comment = await extractComments(i);
        if (comment.username) {
          allComments.push(comment);
          console.log(
            `Extracted comment from page ${pageIdx}, element ${i}:`,
            comment
          );
        }
      } catch (error) {
        console.error(
          `Error extracting comment from page ${pageIdx}, element ${i}:`,
          error
        );
      }
    }
    // Close the current page
    //await page.close();
  }

  // Create the 'Lihkg_data' directory if it doesn't exist
  const dataFolderPath = path.join(__dirname, "Lihkg_data");
  if (!fs.existsSync(dataFolderPath)) {
    fs.mkdirSync(dataFolderPath);
  }

  // Save all comments to a JavaScript file
  //const jsFilePath = path.join(dataFolderPath, `${threadID}_comments.json`);
  //const jsFileContent = `${JSON.stringify(allComments, null, 2)}`;
  //fs.writeFileSync(jsFilePath, jsFileContent);

  // Convert the JSON file to an Excel file
  const excelFilePath = path.join(dataFolderPath, `${threadID}_comments.xlsx`);
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Comments");
  worksheet.addRow(["Username", "Date", "Content"]);
  allComments.forEach((comment) => {
    worksheet.addRow([comment.username, comment.date, comment.content]);
  });
  await workbook.xlsx.writeFile(excelFilePath);
  console.log(`Excel file saved at: ${excelFilePath}`);

  // Close the browser when done
  // Comment out the following line if you want to keep the browser window open after the script finishes
  //await browser.close();
})();
