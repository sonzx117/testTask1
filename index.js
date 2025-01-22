const axios = require('axios');
const XLSX = require('xlsx');
const fs = require('fs');

const url = 'https://go.microsoft.com/fwlink/?LinkID=521962';
(async () => {
    try {
    // Lấy dữ liệu từ URL
    const { data } = await axios.get(url, { responseType: 'arraybuffer' });
    // Đọc dữ liệu từ file excel
    const workbook = XLSX.read(data, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    // Chuyển sheet thành mảng dữ liệu    
    let dataRows = XLSX.utils.sheet_to_json(sheet).map(row =>
      Object.fromEntries(Object.entries(row).map(([key, value]) => [key.trim(), value]))
        );
    // Lọc dữ liệu theo Sales > 50000
    const filteredData = dataRows.filter(row => row['Sales'] > 50000);

    if (filteredData.length === 0) {
      console.log('Không có dữ liệu nào đáp ứng điều kiện lọc.');
      return;
        }
    // Tạo file excel mới từ dữ liệu đã lọc
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(filteredData);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'FilteredData');
    // Xuất file excel    
    const outputFile = 'loc_sales.xlsx';
    XLSX.writeFile(newWorkbook, outputFile);
    // Kiểm tra file đã được tạo thành công
    fs.access(outputFile, fs.constants.F_OK, (err) => {
      if (err) {
        console.error('Lỗi khi kiểm tra file:', err);
        return;
      }
      console.log(`File đã được tạo thành công: ${outputFile}`);
    });

  } catch (error) {
    console.error('Lỗi:', error);
  }
})();
