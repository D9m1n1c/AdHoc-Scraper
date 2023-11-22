const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const targetURL =
  "https://www.discuss.com.hk/viewthread.php?tid=31184058&extra=&page=1"; // You can change this URL
const totalPages = 1; // The total number of pages you want to scrape
const threadID = targetURL.match(/tid=(\d+)/)[1];

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
    const pageURL = targetURL.replace(/page=\d+/, `page=${pageIdx}`);
    await page.goto(pageURL, {
      waitForSelector: "`#mainbox-container-id",
      waitUntil: "domcontentloaded",
    });
    await page.waitForTimeout(5000);

    // Function to extract comments from a specific element index
    const extractComments = async (index) => {
      const usernameSelector = `#mainbox-container-id form div:nth-child(${index}) table tbody tr td:nth-child(1) div.author-detail div.autor-name-row a`;
      const dateSelector = `#mainbox-container-id form div:nth-child(${index}) table tbody tr td:nth-child(2) div.postinfo div.post-date`;
      const contentSelector = `#mainbox-container-id form div:nth-child(${index}) table tbody tr td:nth-child(2) div.postmessage div.postmessage-content span`;
      const username = await page.evaluate((selector) => {
        const element = document.querySelector(selector);
        return element ? element.textContent.trim() : "";
      }, usernameSelector);

      const date = await page.evaluate((selector) => {
        const element = document.querySelector(selector);
        if (element) {
          const text = element.textContent.trim();
          const dateAndTime = text.match(/\d{4}-\d{1,2}-\d{1,2} \d{2}:\d{2}/);
          return dateAndTime ? dateAndTime[0] : "";
        }
        return "";
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
  const dataFolderPath = path.join(__dirname, "HkDiscuss_data");
  if (!fs.existsSync(dataFolderPath)) {
    fs.mkdirSync(dataFolderPath);
  }

  // Save all comments to a JavaScript file
  //const jsFilePath = path.join(dataFolderPath, `${threadID}_comments.json`);
  //const jsFileContent = `${JSON.stringify(allComments, null, 2)}`;
  //fs.writeFileSync(jsFilePath, jsFileContent);

  // Convert the JSON file to an Excel file
  const excelFilePath = path.join(
    dataFolderPath,
    `${threadID}_HkDisucss_comments.xlsx`
  );
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
  await browser.close();
})();
