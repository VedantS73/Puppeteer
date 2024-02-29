const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const fs = require('fs');

// Function to read Excel file and extract ISBN numbers
function readExcel(fileName) {
    const workbook = xlsx.readFile(fileName);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    const isbns = [];

    for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
        const cell = worksheet[xlsx.utils.encode_cell({ r: rowNum, c: 2 })]; // Assuming ISBN column is at index 2 (0-based)
        if (cell && cell.v) {
            isbns.push(cell.v.toString()); // Convert to string since ISBN may contain leading zeros
        }
    }
    return isbns;
}

// Function to scrape data from Snapdeal for a given ISBN
async function scrapeSnapdealData(isbn) {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.goto(`https://www.snapdeal.com/search?keyword=${isbn}`);

    // Clicking on the first product
    await page.waitForSelector('.product-tuple-listing');
    await page.click('.product-tuple-listing');

    console.log('Clicked on the first product');

    
    console.log('Scraping Snapdeal for ISBN:', isbn);

    const titleNode = await page.$('h1'); 

    const title = await page.evaluate(el => el.innerText, titleNode); 

    console.log('Title:', title);

    // const innerText = await page.evaluate(() => {
    //     const h1Element = document.evaluate('/html/body/div[11]/section/div[1]/div[2]/div/div[1]/div[1]/div[1]/h1', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    //     return h1Element ? h1Element.textContent.trim() : null;
    // });

    // console.log('innerText:', innerText);

    // await page.waitForXPath('/html/body/div[11]/section/div[1]/div[2]/div/div[1]/div[1]/div[1]/h1');
    
    // let names = innerText;

    await browser.close();
    return { isbn, title }; // Return the scraped data
}

// Main function to orchestrate the process
async function main() {
    const inputFileName = 'Input.xlsx';
    const outputFileName = 'output.csv';

    // Read ISBNs from Excel
    const isbns = readExcel(inputFileName);

    console.log('Read ISBNs:', isbns);

    // Scrape Snapdeal for each ISBN
    const scrapedData = [];
    for (const isbn of isbns) {
        const data = await scrapeSnapdealData(isbn);
        scrapedData.push(data);
    }

    // // Write scraped data to CSV
    // const csvData = scrapedData.map(({ isbn, bookName }) => `${isbn},${bookName}`).join('\n');
    // fs.writeFileSync(outputFileName, csvData);

    // console.log('Scraping and writing completed successfully.');
}

main().catch(error => console.error(error));