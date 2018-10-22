const Excel = require('./node_modules/exceljs');
const moment = require('moment');

const workingFile = './track.xlsx';

const workbook = new Excel.Workbook();

const outputObject = {};

workbook.xlsx.readFile(workingFile)
.then( output => {

	const worksheet = workbook.getWorksheet("Sheet1");

	worksheet.eachRow(function(row, rowNumber) {
		const thisRow = row.values.slice(1);
		let problem = "";

		const formNumberColumn = thisRow[0];
		const dayColumn = thisRow[2];
		const monthColumn = thisRow[3];
		const yearColumn = thisRow[4];

		let formNumber = "";
		let outputCell = worksheet.getCell(`H${rowNumber}`)

		if (formNumberColumn.length == 31){
			formNumber = formNumberColumn.slice(3, 14);
		}
		else if (formNumberColumn.length == 1){
			formNumber = formNumberColumn;
		}
		else {
			problem = "Something is wrong with the form number"
		}


		const workingDate = moment(`${thisRow[2]}-${thisRow[3]}
-${thisRow[4]}`, "DD-MM-YYYY");
		console.log({workingDate});
		let dayOB = workingDate.format("DD");
		console.log(typeof(dayOB));
		console.log({thisRow});

		if (workingDate.isValid() === false){
			problem = "Something is wrong with the date of birth."
		}

		if (problem == true){
			outputObject[outputCell] = problem;

		}
	})
}
);

console.log(outputObject);

// //Update a cell
// row.getCell(1).value = 5;
//
// row.commit();
//
// //Save the workbook
// return workbook.xlsx.writeFile("data/Sample.xlsx");
