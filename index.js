const Excel = require('./node_modules/exceljs');

const workingFile = './track.xlsx';

const workbook = new Excel.Workbook();

workbook.xlsx.readFile(workingFile)
.then( output => {

	const worksheet = workbook.getWorksheet("Sheet1");
	worksheet.eachRow(function(row, rowNumber) {
		const thisRow = row.values.slice(1);
		let problem = "";
		let dayCell = thisRow[3];
		let monthCell = thisRow[4];

		if(dayCell <= 31 && dayCell > 0){
			if(dayCell < 10){
				thisRow[3] = `0${dayCell.toString()}`;
				console.log({dayCell});
			}
		}
		else {
			problem = `Day is outside valid range in row ${row}`;
		}

		if(monthCell <= 12 && monthCell > 0){
			if(monthCell < 10){
				monthCell = "0" + monthCell.toString();
			}
		}
		else {
			problem = `Month is outside valid range in row ${row}`;
		}

		console.log({thisRow});
	})
}
);
