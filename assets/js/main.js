const app = Vue.createApp({
	data() {
		return {
			hello: 'hello123123',
			files: [],
			workbooks: {},
			isParsingInProgress: false,
			finishedFiles: [],
			parsedData: {},
			fioColumnName: 'ФИО сиделки',
			firstSheetName: 'Сводная таблица часов',
			secondSheetName: 'Сводная таблица краткая',
			firthSheetName: 'Часы главных специалистов',
		}
	},
	methods: {
		async filesSelectedHandler(e) {
			this.isParsingInProgress = true;
			this.parsedData = {};
			this.finishedFiles = [];
			this.files = Array.from(e.target.files);
			for (const f of this.files) {
				const wb = await XLSX.read(await f.arrayBuffer());
				const parsedData = this.parseWorkbook(wb);
				if (!parsedData) {
					alert(`Файл ${f.name} не похож на табель. Обработка остановлена`);
					return;
				}
				this.parsedData[f.name.split('.')[0]] = parsedData;
				this.finishedFiles.push(f.name);
			}
			this.isParsingInProgress = false;
		},
		parseWorkbook(wb) {
			const sheet = wb.Sheets[wb.SheetNames[0]];
			if (!this.isHoursWorksheet(sheet)) {
				return false;
			}
			const res = {};
			let row = 16;
			while (1) {
				let fio = sheet[`F${row}`]?.v;
				if (fio === undefined) {
					break;
				}
				let hours = sheet[`FV${row}`]?.v;
				res[fio] = (res[fio] || 0) +  hours;
				row += 2;
				if (row > 500) {
					alert('Что-то пошло не так. Не удалось найти нижнюю границу таблицы.');
					break;
				}
			}
			return res;
		},
		isHoursWorksheet(sheet) {
			if (sheet['CN6']?.v !== 'ТАБЕЛЬ') {
				return false;
			}
			return true;
		},
		async exportToXlsx() {
			const wb = new ExcelJS.Workbook();
			this.addFirstWorksheet(wb);
			this.addSecondWorksheet(wb);
			this.addFirthWorksheet(wb);

			const buffer = await wb.xlsx.writeBuffer();
			this.downloadBuffer(buffer);
		},
		downloadBuffer(buffer) {
			var blob = new Blob([buffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
			var link = document.createElement('a');
			link.href = window.URL.createObjectURL(blob);
			link.download = 'Сводная таблица.xlsx';
			link.click();
		},
		addFirstWorksheet(wb) {
			const worksheet = wb.addWorksheet(this.firstSheetName);
			let columnNames = [ this.fioColumnName, ...this.bosses, 'Сумма' ];
			worksheet.columns = columnNames;
			const headerRow = worksheet.addRow(columnNames);
			const maxBossNameLength = this.bosses.reduce((acc, boss) => Math.max(boss.length, acc), 0);
			const charWidth = 7;
			// высота строки, чтобы можно было повернуть имена начальников вертикально
			headerRow.height = maxBossNameLength * charWidth;

			// ширина ФИО сиделки
			worksheet.getColumn(1).width = this.maxNurseNameLength * 1.1;
			// поворот фамилий начальников вертикально
			for (let i = 2; i < this.bosses.length + 2; i++) {
				headerRow.getCell(i).style = { alignment: { textRotation: 90, horizontal: 'center' } };
				worksheet.getColumn(i).width = 6;
			}

			// суммы по начальникам
			const sumRow = [`Количество сиделок: ${this.resultData.length}`];
			const nurseCount = this.resultData.length;
			const mainDataRowsStartAt = 3;
			for (let i = 1; i <= this.bosses.length + 1; i++) {
				const colLetter = this.convertIndexToLetter(i);
				sumRow.push({
					formula: `SUM(${colLetter}${mainDataRowsStartAt}:${colLetter}${mainDataRowsStartAt + nurseCount - 1})`
				})
			}
			worksheet.addRow(sumRow);

			// вставка основных данных
			const fromLetter = this.convertIndexToLetter(1);
			const toLetter = this.convertIndexToLetter(this.bosses.length);
			this.resultData.forEach((row, idx) => {
				const rowNumber = idx + 3;
				const values = [
					row[this.fioColumnName],
					...this.bosses.map(b => row[b]),
					{formula: `SUM(${fromLetter}${rowNumber}:${toLetter}${rowNumber})`}
				];
				worksheet.addRow(values);
			})
		},
		addSecondWorksheet(wb) {
			const worksheet = wb.addWorksheet(this.secondSheetName);
			worksheet.getColumn(1).width = this.maxNurseNameLength * 1.1;
			const sumColumnName = this.convertIndexToLetter(1 + this.bossesCount);
			this.resultData.forEach((nurse, idx) => {
				worksheet.addRow([
					{ formula: `'${this.firstSheetName}'!A${idx + 3}`},
					{ formula: `'${this.firstSheetName}'!${sumColumnName}${idx + 3}`},
				])
			})
		},
		addFirthWorksheet(wb) {
			const worksheet = wb.addWorksheet(this.firthSheetName);
			worksheet.getColumn(2).width = 30;
			worksheet.getColumn(3).width = 10;
			worksheet.getColumn(4).width = 15;
			worksheet.getColumn(5).width = 20;
			worksheet.addRow(['', '', 'Зарплата', '', '']);
			const header = ['№ п/п', 'ФИО ГЛ.СПЕЦИАЛИСТА', 'ТАБЕЛЬ СДАН', 'ПРОВЕРЕН, ЧАСЫ', 'ЭЛЕКТР.ВАРИАНТ, ЧАСЫ'];
			const headerRow = worksheet.addRow(header);
			headerRow.height = 30;
			headerRow.eachCell(c => {
				this.addBorderToCell(c);
				c.style = {
					alignment: { wrapText: true, horizontal: 'center' },
					font: { bold: true },
				}
			});
			for (let i = 1; i <= this.bossesCount; i++) {
				const bossColumnName = this.convertIndexToLetter(i);
				const rowData = [
					i,
					{ formula: `'${this.firstSheetName}'!${bossColumnName}1` },
					'',
					'',
					{ formula: `'${this.firstSheetName}'!${bossColumnName}2` },
				];
				worksheet.addRow(rowData).eachCell(this.addBorderToCell);
			}
		},
		addBorderToCell(cell) {
			cell.border = {
				top: { style: 'thin' },
				left: { style: 'thin' },
				bottom: { style: 'thin' },
				right: { style: 'thin' }
			};
		},
		convertIndexToLetter(index) {
			const result = [];
			do  {
				const reminder = index % 26;
				result.unshift(String.fromCharCode(reminder + 65));
				index = Math.floor(index / 26);
			} while (index > 0)
			return result.join();
		}
	},
	// mounted: {},
	watch: {
		files: function() {
			// console.log(this.files);
		},
	},
	computed: {
		filesCount() {
			return this.files.length;
		},
		fileNames() {
			return this.files.map(f => f.name);
		},
		bosses() {
			return this.fileNames.map(f => f.split('.')[0]);
		},
		areAllFilesParsed() {
			return this.filesCount > 0 && this.filesCount === this.finishedFiles.length;
		},
		maxNurseNameLength() {
			return this.resultData.reduce((acc, row) => Math.max(row[this.fioColumnName].length, acc), 0);
		},
		bossesCount() {
			return this.bosses.length;
		},
		resultData() {
			const res = {};
			Object.entries(this.parsedData).forEach(([boss, nurses]) => {
				Object.entries(nurses).forEach(([nurse, hours]) => {
					res[nurse] = {
						...(res[nurse] || {}),
						[this.fioColumnName]: nurse,
						[boss]: (res[nurse]?.[boss] || 0) + hours,
					}
				})
			})
			return Object.values(res);
		},
		resultWorkbook() {
			const bosses  = this.fileNames.map(n => n.split('.')[0]);
			const header = [this.fioColumnName, ...bosses, 'Сумма'];

			const wb = XLSX.utils.book_new();
			const ws = XLSX.utils.aoa_to_sheet([header]);
			ws['A1'].s = {bold: true};
			XLSX.utils.book_append_sheet(wb, ws, 'Сводная таблица часов');

			return wb;

			Object.keys(this.resultData).forEach(boss => {
				header.push(boss.split('.')[0]);
			})
			header.push('Сумма');
			XLSX.to_workbook
			const rawData = [];
			Object.values(this.parsedData).forEach(nurses => {
				Object.entries(this.parsedData).forEach(([nurse, hours]) => {

				})
			})
		}
	},
});

// Vue.createApp(app);
// app.config.globalProperties.dayjs = dayjs;
app.mount('#app')
