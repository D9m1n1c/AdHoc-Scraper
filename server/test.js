const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// Define an array of target URLs
const targetURLs = [
  "https://www.facebook.com/AAStocks.com.Limited/posts/666952455467559",
  "https://www.facebook.com/AAStocks.com.Limited/posts/666910578805080",
  "https://www.facebook.com/inmediahknet/posts/668320112010469",
  "https://www.facebook.com/now.comNews/posts/267378232716329",
  "https://www.facebook.com/HongKongGoodNews/posts/691945706309093",
  "https://www.facebook.com/hketpage/posts/693584086146214",
  "https://www.facebook.com/MoneyFlowHK/posts/774258178041277",
  "https://www.facebook.com/etnet.com.hk/posts/674389994724493",
  "https://www.facebook.com/hongkongeconomicjournal/posts/672275568266580",
  "https://www.facebook.com/bossmindmedia/posts/323747566668443",
  "https://www.facebook.com/FaSecrets/posts/303310035593473",
  "https://www.facebook.com/bbwhk/posts/314342657647746",
  "https://www.facebook.com/tvbnewsofficial/posts/157939187323564",
  "https://www.facebook.com/imoneymagazine/posts/687880660049306",
  "https://www.facebook.com/CapitalPlatformHK/posts/701873641956456",
  "https://www.facebook.com/now.comFinance/posts/684599887022685",
  "https://www.facebook.com/hkafofficial/posts/782677703860549",
  "https://www.facebook.com/tkp1902financialnews/posts/250568454572257",
];

(async () => {
  // Launch a non-headless browser
  const browser = await puppeteer.launch({
    executablePath:
      "C:/Users/cchue/.cache/puppeteer/chrome/win64-116.0.5845.96/chrome-win64/chrome.exe",
    headless: false,
    defaultViewport: null,
  });

  for (const targetURL of targetURLs) {
    try {
      const page = await browser.newPage();

      let threadID;

      // Extract thread ID from the URL
      const urlSections = targetURL.split("/");
      const lastSection = urlSections[urlSections.length - 1];
      threadID = `${lastSection}`;
      // if (match) {
      //   threadID = match[1];
      // } else {
      //   const urlSections = targetURL.split("/");
      //   const lastSection = urlSections[urlSections.length - 1];
      //   threadID = `${lastSection}_CX880`;
      // }

      // Navigate to the target page
      await page.goto(targetURL, {
        waitUntil: "domcontentloaded",
      });

      await page.waitForTimeout(3000);

      // Handle cookies button
      try {
        await page.click(
          "body > div.__fb-light-mode.x1n2onr6.x1vjfegm > div.x9f619.x1n2onr6.x1ja2u2z > div > div.x1uvtmcs.x4k7w5x.x1h91t0o.x1beo9mf.xaigb6o.x12ejxvf.x3igimt.xarpa2k.xedcshv.x1lytzrv.x1t2pt76.x7ja8zs.x1n2onr6.x1qrby5j.x1jfb8zj > div > div > div > div.xzg4506.x1l90r2v.x1pi30zi.x1swvt13 > div > div:nth-child(2) > div.x1i10hfl.xjbqb8w.x6umtig.x1b1mbwd.xaqea5y.xav7gou.x1ypdohk.xe8uvvx.xdj266r.x11i5rnm.xat24cr.x1mh8g0r.xexx8yu.x4uap5.x18d9i69.xkhd6sd.x16tdsg8.x1hl2dhg.xggy1nq.x1o1ewxj.x3x9cwd.x1e5q0jg.x13rtm0m.x87ps6o.x1lku1pv.x1a2a7pz.x9f619.x3nfvp2.xdt5ytf.xl56j7k.x1n2onr6.xh8yej3 > div > div.x6s0dn4.x78zum5.xl56j7k.x1608yet.xljgi0e.x1e0frkt > div > span > span"
        );
        console.log("Click cookie button");
      } catch (error) {
        console.log("No cookies button");
      }

      // Scroll function
      const scrollAndLoadComments = async () => {
        let previousHeight = 0;

        while (true) {
          // Scroll to the bottom
          await page.evaluate("window.scrollTo(0, document.body.scrollHeight)");

          // Wait for some time to allow content to load
          await page.waitForTimeout(1000); // Adjust the timeout as needed

          // Get the new scroll height after scrolling
          const newHeight = await page.evaluate("document.body.scrollHeight");

          // If the scroll height hasn't changed, there's no more content to load
          if (newHeight === previousHeight) {
            break; // Exit the loop
          }

          // Update the previous scroll height
          previousHeight = newHeight;
        }
      };

      // Click on the initial element to load comments
      try {
        await page.click(
          "div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div.x78zum5.xdt5ytf.x1t2pt76.x1n2onr6.x1ja2u2z.x10cihs4 > div.x1qjc9v5.x78zum5.xl56j7k.x193iq5w.x1t2pt76 > div > div > div > div > div > div > div > div > div > div > div > div > div > div:nth-child(2) > div > div > div:nth-child(4) > div > div > div.x2bj2ny.x12nagc > div:nth-child(3) > div > div > div > span"
        );
        await page.waitForTimeout(2000);

        // Click on the "View all comments" button
        await page.click(
          "div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div:nth-child(2) > div > div > div.xu96u03.xm80bdy.x10l6tqk.x13vifvy > div.x1uvtmcs.x4k7w5x.x1h91t0o.x1beo9mf.xaigb6o.x12ejxvf.x3igimt.xarpa2k.xedcshv.x1lytzrv.x1t2pt76.x7ja8zs.x1n2onr6.x1qrby5j.x1jfb8zj > div > div > div > div > div > div > div.x78zum5.xdt5ytf.x1iyjqo2.x1n2onr6 > div > div:nth-child(3) > div.x6s0dn4.x78zum5.x1q0g3np.x1iyjqo2.x1qughib.xeuugli > div > div:nth-child(1) > span"
        );
        await page.waitForTimeout(2000);

        await clickViewMore();
        await page.waitForTimeout(2000);
        await scrollAndLoadComments();
      } catch (error) {
        console.log("====", error);
        continue;
      }

      async function clickViewMore() {
        while (true) {
          const viewMoreButton = await page.$(
            "div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div.x78zum5.xdt5ytf.x1t2pt76.x1n2onr6.x1ja2u2z.x10cihs4 > div.x1qjc9v5.x78zum5.xl56j7k.x193iq5w.x1t2pt76 > div > div > div > div > div > div > div > div > div > div > div > div > div > div:nth-child(2) > div > div > div:nth-child(4) > div > div > div.x2bj2ny.x12nagc > div:nth-child(5) > div.x78zum5.x1iyjqo2.x21xpn4.x1n2onr6 > div.x1i10hfl.xjbqb8w.xjqpnuy.xa49m3k.xqeqjp1.x2hbi6w.x13fuv20.xu3j5b3.x1q0q8m5.x26u7qi.x972fbf.xcfux6l.x1qhh985.xm0m39n.x9f619.x1ypdohk.xdl72j9.xe8uvvx.xdj266r.x11i5rnm.xat24cr.x1mh8g0r.x2lwn1j.xeuugli.xexx8yu.x18d9i69.xkhd6sd.x1n2onr6.x16tdsg8.x1hl2dhg.xggy1nq.x1ja2u2z.x1t137rt.x1o1ewxj.x3x9cwd.x1e5q0jg.x13rtm0m.x3nfvp2.x1q0g3np.x87ps6o.x1a2a7pz.x6s0dn4.xi81zsa.x1iyjqo2.xs83m0k.xsyo7zv.xt0b8zv > span"
          );

          if (!viewMoreButton) {
            break;
          }
          await viewMoreButton.click();
          await page.waitForTimeout(2000);
        }
      }

      const allComments = [];

      // Extract comments
      const comments = await page.$$(
        "div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div.x78zum5.xdt5ytf.x1t2pt76.x1n2onr6.x1ja2u2z.x10cihs4 > div.x1qjc9v5.x78zum5.xl56j7k.x193iq5w.x1t2pt76 > div > div > div > div > div > div > div > div > div > div > div > div > div > div:nth-child(2) > div > div > div:nth-child(4) > div > div > div.x2bj2ny.x12nagc > ul > li"
      );
      for (let i = 1; i <= comments.length; i++) {
        const comment = await extractComments(i);
        if (comment) {
          allComments.push(comment);
          console.log(`Extracted comment ${i}:`, comment);
        }
      }

      async function extractComments(index) {
        const usernameSelector = `div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div.x78zum5.xdt5ytf.x1t2pt76.x1n2onr6.x1ja2u2z.x10cihs4 > div.x1qjc9v5.x78zum5.xl56j7k.x193iq5w.x1t2pt76 > div > div > div > div > div > div > div > div > div > div > div > div > div > div:nth-child(2) > div > div > div:nth-child(4) > div > div > div.x2bj2ny.x12nagc > ul > li:nth-child(${index}) > div.x1n2onr6.x1iorvi4.x4uap5.x18d9i69.x1swvt13.x78zum5.x1q0g3np.x1a2a7pz > div.x1r8uery.x1iyjqo2.x6ikm8r.x10wlt62.x1pi30zi > div:nth-child(1) > div.xv55zj0.x1vvkbs.x1rg5ohu.xxymvpz > div > div.xdl72j9.x1iyjqo2.xs83m0k.xeuugli.xh8yej3 > div > div > span > span`;
        const dateSelector = ` div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div.x78zum5.xdt5ytf.x1t2pt76.x1n2onr6.x1ja2u2z.x10cihs4 > div.x1qjc9v5.x78zum5.xl56j7k.x193iq5w.x1t2pt76 > div > div > div > div > div > div > div > div > div > div > div > div > div > div:nth-child(2) > div > div > div:nth-child(4) > div > div > div.x2bj2ny.x12nagc > ul > li:nth-child(${index}) > div.x1n2onr6.x1iorvi4.x4uap5.x18d9i69.x1swvt13.x78zum5.x1q0g3np.x1a2a7pz > div.x1r8uery.x1iyjqo2.x6ikm8r.x10wlt62.x1pi30zi > div.x6s0dn4.x3nfvp2 > ul > li > div > a > span`;
        const contentSelector = `div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div.x78zum5.xdt5ytf.x1t2pt76.x1n2onr6.x1ja2u2z.x10cihs4 > div.x1qjc9v5.x78zum5.xl56j7k.x193iq5w.x1t2pt76 > div > div > div > div > div > div > div > div > div > div > div > div > div > div:nth-child(2) > div > div > div:nth-child(4) > div > div > div.x2bj2ny.x12nagc > ul > li:nth-child(${index}) > div.x1n2onr6.x4uap5.x18d9i69.x1swvt13.x1iorvi4.x78zum5.x1q0g3np.x1a2a7pz > div.x1r8uery.x1iyjqo2.x6ikm8r.x10wlt62.x1pi30zi > div:nth-child(1) > div.xv55zj0.x1vvkbs.x1rg5ohu.xxymvpz > div > div.xdl72j9.x1iyjqo2.xs83m0k.xeuugli.xh8yej3 > div > div > div > span > div > div`; // Add your content selector here

        const username = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "";
        }, usernameSelector);

        const date = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "";
        }, dateSelector);

        const content = await page.evaluate((selector) => {
          const element = document.querySelector(selector);
          return element ? element.textContent.trim() : "";
        }, contentSelector);

        if (!username || username === "") {
          const usernameSelector2 = `div > div:nth-child(1) > div > div.x9f619.x1n2onr6.x1ja2u2z > div > div > div > div.x78zum5.xdt5ytf.x1t2pt76.x1n2onr6.x1ja2u2z.x10cihs4 > div.x1qjc9v5.x78zum5.xl56j7k.x193iq5w.x1t2pt76 > div > div > div > div > div > div > div > div > div > div > div > div > div > div:nth-child(2) > div > div > div:nth-child(4) > div > div > div.x2bj2ny.x12nagc > ul > li:nth-child(${index}) > div.x1n2onr6.x4uap5.x18d9i69.x1swvt13.x1iorvi4.x78zum5.x1q0g3np.x1a2a7pz > div.x1r8uery.x1iyjqo2.x6ikm8r.x10wlt62.x1pi30zi > div:nth-child(1) > div.xv55zj0.x1vvkbs.x1rg5ohu.xxymvpz > div > div.xdl72j9.x1iyjqo2.xs83m0k.xeuugli.xh8yej3 > div > div > span > a > span > span`; // Replace 'xxx' with your path 2 selector
          const username2 = await page.evaluate((selector) => {
            const element = document.querySelector(selector);
            return element ? element.textContent.trim() : "";
          }, usernameSelector2);
          return { date, username: username2, content };
        }
        return { date, username, content };
      }

      // Create the 'Facebook_data' directory if it doesn't exist
      const dataFolderPath = path.join(__dirname, "CX880_FB_data");
      if (!fs.existsSync(dataFolderPath)) {
        fs.mkdirSync(dataFolderPath);
      }

      // Save all comments to an Excel file
      const excelFilePath = path.join(
        dataFolderPath,
        `${threadID}_CX880_FB_comments.xlsx`
      );
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Comments");
      worksheet.addRow(["Date", "Username", "Content"]);
      allComments.forEach((comment) => {
        worksheet.addRow([comment.date, comment.username, comment.content]);
      });

      await workbook.xlsx.writeFile(excelFilePath);
      console.log(`Excel file saved at: ${excelFilePath}`);

      // Close the current page
      await page.close();
    } catch (error) {
      continue;
    }
  }

  // Close the browser when done
  // Comment out the following line if you want to keep the browser window open after the script finishes
  await browser.close();
})();
