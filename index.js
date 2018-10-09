const Excel = require('./node_modules/exceljs');
const moment = require('moment');

const workingFile = './track.xlsx';

const workbook = new Excel.Workbook();

workbook.xlsx.readFile(workingFile)
.then( output => {

	const worksheet = workbook.getWorksheet("Sheet1");
	worksheet.eachRow(function(row, rowNumber) {
		const thisRow = row.values.slice(1);
		let problem = "";

		const workingDate = moment(`${thisRow[2]}-${thisRow[3]}
-${thisRow[4]}`, "DD-MM-YYYY");
		console.log({workingDate});
		let dayOB = workingDate.format("DD");
		console.log(typeof(dayOB));
		console.log({thisRow});
	})
}
);
