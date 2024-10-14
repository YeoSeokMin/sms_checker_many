document.getElementById('processFiles').addEventListener('click', async () => {
    const input = document.getElementById('fileInput');
    const files = input.files;

    if (files.length === 0) {
        alert('Please select files.');
        return;
    }

    document.getElementById('status').textContent = 'Processing...';
    const progressBar = document.getElementById('progress');
    const progressValue = document.getElementById('progressValue');

    let allNumbers = [];
    let totalFiles = files.length;
    let processedFiles = 0;

    for (const file of files) {
        const data = await readFile(file);
        const numbers = extractNumbersFromExcel(data);
        allNumbers = allNumbers.concat(numbers);

        // 파일 처리 후 진행률 업데이트
        processedFiles++;
        const progressPercentage = Math.round((processedFiles / totalFiles) * 100);
        progressBar.value = progressPercentage;
        progressValue.textContent = `${progressPercentage}%`;
    }

    if (allNumbers.length === 0) {
        document.getElementById('status').textContent = 'No numeric data found.';
        return;
    }

    splitAndDownload(allNumbers);
    document.getElementById('status').textContent = 'Files ready for download.';
    progressBar.value = 100;
    progressValue.textContent = '100%';
});

// 엑셀 파일을 읽는 함수
function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            resolve(workbook);
        };
        reader.readAsArrayBuffer(file);
    });
}

// 엑셀에서 숫자 데이터만 추출
function extractNumbersFromExcel(workbook) {
    let numbers = [];

    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach(row => {
            row.forEach(cell => {
                // 숫자만 배열에 추가
                if (typeof cell === 'number') {
                    numbers.push(cell);
                }
            });
        });
    });

    return numbers;
}

// 숫자 데이터를 50만 개씩 나눠 엑셀로 다운로드
function splitAndDownload(numbers) {
    const batchSize = 500000;
    let fileCount = 1;

    for (let i = 0; i < numbers.length; i += batchSize) {
        const chunk = numbers.slice(i, i + batchSize);
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(chunk.map(num => [num]));

        XLSX.utils.book_append_sheet(wb, ws, 'Numbers');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = `numbers_part_${fileCount}.xlsx`;
        a.click();

        URL.revokeObjectURL(url);
        fileCount++;
    }
}
