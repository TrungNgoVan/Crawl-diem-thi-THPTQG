const puppeteer = require('puppeteer');
const Excel = require('exceljs');

(async () => {
    // Load workbook and worksheet
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile('SBD.xlsx');
    const worksheet = workbook.getWorksheet('Sheet1');

    // Get SBD values
    const sbdColumn = worksheet.getColumn('A');
    const sbdValues = sbdColumn.values; // Ignore header row

    // Start 
    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    // Create worksheet excel
    const resultWorkbook = new Excel.Workbook();
    const resultWorksheet = resultWorkbook.addWorksheet('Dữ liệu');

    // Add table
    resultWorksheet.columns = [
        { header: 'SBD', key: 'sbd' },
        { header: 'Toán', key: 'toan' },
        { header: 'Ngữ văn', key: 'nguvan' },
        { header: 'Ngoại ngữ', key: 'ngoaingu' },
        { header: 'Vật lý', key: 'vatly' },
        { header: 'Hóa học', key: 'hoahoc' },
        { header: 'Sinh học', key: 'sinhhoc' },
        { header: 'Giáo dục công dân', key: 'gdcd' },
        { header: 'Lịch sử', key: 'lichsu' },
        { header: 'Địa lý', key: 'dialy' },
    ];
    var temp;
    for (const termSbd of sbdValues) {
        try {
            const sbd = termSbd.toString();
            temp = sbd;
            console.log(`Lấy dữ liệu cho SBD: ${sbd}`);

            // Enter SBD
            await page.goto('https://diemthi.vnexpress.net/', { waitUntil: 'networkidle2' });
            await page.type('#keyword', sbd);

            // Press submit
            await Promise.all([
                page.waitForNavigation({ waitUntil: 'networkidle2' }),
                page.click('.search_submit'),
            ]);

            // Waiting e-table upload
            await page.waitForSelector('.e-table', { timeout: 10000 });

            // Exploit data
            const data = await page.evaluate(() => {
                const table = document.querySelector('.e-table');
                const rows = Array.from(table.querySelectorAll('tr'));
                return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll('td'));
                    return cells.map(cell => cell.textContent.trim());
                });
            });

            // filter data
            const filteredData = data.slice(1, -1);
            console.log(filteredData);

            // Add data to worksheet
            const rowData = {
                sbd: sbd,
                toan: '',
                nguvan: '',
                ngoaingu: '',
                vatly: '',
                hoahoc: '',
                sinhhoc: '',
                gdcd: '',
                lichsu: '',
                dialy: '',
            };
            filteredData.forEach((row, index) => {
                row.forEach((cell, index) => {
                    if (cell === 'Toán') {
                        rowData.toan = row[index + 1];
                    } else if (cell === 'Ngữ văn') {
                        rowData.nguvan = row[index + 1];
                    } else if (cell === 'Ngoại ngữ') {
                        rowData.ngoaingu = row[index + 1];
                    } else if (cell === 'Vật lý') {
                        rowData.vatly = row[index + 1];
                    } else if (cell === 'Hóa học') {
                        rowData.hoahoc = row[index + 1];
                    } else if (cell === 'Sinh học') {
                        rowData.sinhhoc = row[index + 1];
                    } else if (cell === 'Giáo dục công dân') {
                        rowData.gdcd = row[index + 1];
                    } else if (cell === 'Lịch sử') {
                        rowData.lichsu = row[index + 1];
                    } else if (cell === 'Địa lý') {
                        rowData.dialy = row[index + 1];
                    }
                });

            });
            resultWorksheet.addRow(rowData);
        } catch (error) {
            console.log('Error');
            resultWorksheet.addRow({ sbd: temp });
            continue;
        }
    }

    // Export file excel
    await resultWorkbook.xlsx.writeFile('finalData.xlsx');

    // Close browser
    await browser.close();
})();
