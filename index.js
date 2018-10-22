const Excel = require('./node_modules/exceljs');
const moment = require('moment');
const puppeteer = require('puppeteer');

const workingFile = './track.xlsx';

const workbook = new Excel.Workbook();

let outputObject = {};

async function read(){
workbook.xlsx.readFile(workingFile)
.then( await function(output){

	const worksheet = workbook.getWorksheet("Sheet1");

	worksheet.eachRow(function(row, rowNumber) {
		const thisRow = row.values.slice(1);
		let problem = "";

		const formNumberColumn = thisRow[0];
		const dayColumn = thisRow[1];
		const monthColumn = thisRow[2];
		const yearColumn = thisRow[3];

		let formNumber = "";
		let outputCell = worksheet.getCell(`H${rowNumber}`)

		if (formNumberColumn.length == 31){
			formNumber = formNumberColumn.slice(3, 14);
		}
		else if (formNumberColumn.length == 13){
			formNumber = formNumberColumn;
		}
		else {
			problem = "Something is wrong with the form number"
		}


		const workingDate = moment(`${dayColumn}-${monthColumn}
-${yearColumn}`, "DD-MM-YYYY");
		let dayOB = workingDate.format("DD");
		let monthOB = workingDate.format("MM");
		let yearOB = workingDate.format("YYYY");


		if (workingDate.isValid() === false){
			problem = "Something is wrong with the date of birth."
		}

		if (problem !== ""){
			outputObject[`H${rowNumber}`] = problem;
			console.log(outputObject)
		} else { async function puppet(){
			console.log(("Puppet"));

			let browser = await puppeteer.launch();
			page = await browser.newPage;
			await page.goto("https://secure.crbonline.gov.uk/enquiry/enquirySearchAction.do")
			await page.click('[id="AppNo"]');
			await page.type('[id="AppNo"]', formNumber)
			await page.click('[id="dateOfBirthDay"]');
			await page.select('[id="dateOfBirthDay"]', dayOB);
			await page.click('[id="dateOfBirthMonth"]');
			await page.select('[id="dateOfBirthMonth"]', monthOB);
			await page.click('[id="dateOfBirthYear"]');
			await page.select('[id="dateOfBirthYear"]', yearOB);
			const output = await page.click('[name="submit"]');

			console.log(output);
		}


		}
	})
}
);
}


read().
then( output => console.log(outputObject)).
catch(err => console.log(err))

// //Update a cell
// row.getCell(1).value = 5;
//
// row.commit();
//
// //Save the workbook
// return workbook.xlsx.writeFile("data/Sample.xlsx");
