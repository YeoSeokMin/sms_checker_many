var phoneNumbers = [];
var uniquePhoneNumbers = [];
var invalidCount = 0;
var totalCount = 0;
var removedDuplicateCount = 0;
var emptyCount = 0;
var fileName = '';

$(document).ready(function () {
    $('#fileUpload').on('change', handleFileUpload);
    $('#confirmButton').on('click', displayResults);
    $('#toggleResultButton').on('click', () => $('#result').toggle());
    $('#rangeCopyButton').on('click', handleRangeCopy);
});

async function handleFileUpload(e) {
    resetState();
    disableConfirmButton();

    const files = Array.from(e.target.files);
    fileName = getFileNameWithoutExtension(files[0].name);

    console.clear();
    console.log(`파일 업로드 시작: ${files.length}개의 파일`);
    
    let processedFiles = 0;
    const totalFiles = files.length;

    await Promise.all(files.map(file => {
        return readFile(file).then(() => {
            processedFiles++;
            updateProgressBar(processedFiles, totalFiles);
        });
    }));

    processPhoneNumbers();
    enableConfirmButton();
    $('#confirmButton').trigger('click');
}

function updateProgressBar(processedFiles, totalFiles) {
    const progress = (processedFiles / totalFiles) * 100;
    $('#progressBar').css('width', progress + '%');
    $('#progressText').text(Math.round(progress) + '%');
}

function resetState() {
    phoneNumbers = [];
    uniquePhoneNumbers = [];
    invalidCount = 0;
    totalCount = 0;
    removedDuplicateCount = 0;
    emptyCount = 0;
}

function disableConfirmButton() {
    $('#confirmButton').prop('disabled', true);
}

function enableConfirmButton() {
    $('#confirmButton').prop('disabled', false);
}

function getFileNameWithoutExtension(fileName) {
    return fileName.split('.').slice(0, -1).join('.');
}

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                if (!workbook.SheetNames.length) {
                    console.warn(`빈 파일입니다: ${file.name}`);
                    resolve();
                    return;
                }
                handleWorkbook(workbook, file.name);
                resolve();
            } catch (error) {
                console.error('파일을 읽는 중 오류가 발생했습니다.', error);
                resolve();
            }
        };
        reader.onerror = () => {
            console.error(`파일을 읽는 중 오류가 발생했습니다: ${file.name}`);
            resolve();
        };
        reader.readAsArrayBuffer(file);
    });
}

function handleWorkbook(workbook, fileName) {
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet || !worksheet['!ref']) {
            console.warn(`비어 있는 시트입니다: ${sheetName}`);
            return;
        }
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const columnData = getColumnDataWithMaxMatches(worksheet, range);
        if (columnData.targetColumn >= 0) {
            const { totalCount, emptyCount } = extractPhoneNumbersFromColumn(worksheet, range, columnData.targetColumn);
            const columnLetter = XLSX.utils.encode_col(columnData.targetColumn); // 열 알파벳 추적
            console.log(`파일 처리 완료: ${fileName} (시트 개수: ${workbook.SheetNames.length})(추적 열: ${columnLetter}열, 데이터 총 개수: ${totalCount}, 공란 개수: ${emptyCount})`);
        }
    });
}

function getColumnDataWithMaxMatches(worksheet, range) {
    const regexPatterns = [
        /010[- ]?\d{4}[- ]?\d{4}/,
        /010\d{7,8}/,
        /82\d{2}[- ]?\d{3,4}[- ]?\d{4}/,
        /82\d{2}\d{7,8}/,
        /01[1-9][- ]?\d{3,4}[- ]?\d{4}/,
        /10[- ]?\d{4}[- ]?\d{4}/,
    ];

    let maxMatchCount = 0;
    let targetColumn = -1;

    for (let C = range.s.c; C <= range.e.c; ++C) {
        let columnMatchCount = 0;
        let matchedNumbers = new Set();

        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cell_address = { c: C, r: R };
            const cell_ref = XLSX.utils.encode_cell(cell_address);
            const cell = worksheet[cell_ref];
            if (cell && cell.v) {
                const cellValue = String(cell.v);
                regexPatterns.forEach(pattern => {
                    const matches = cellValue.match(pattern);
                    if (matches) {
                        matches.forEach(match => matchedNumbers.add(match));
                    }
                });
            }
        }

        columnMatchCount = matchedNumbers.size;

        if (columnMatchCount > maxMatchCount) {
            maxMatchCount = columnMatchCount;
            targetColumn = C;
        }
    }

    return { maxMatchCount, targetColumn };
}

function extractPhoneNumbersFromColumn(worksheet, range, targetColumn) {
    let localTotalCount = 0;
    let localEmptyCount = 0;

    for (let R = range.s.r; R <= range.e.r; ++R) {
        const cell_address = { c: targetColumn, r: R };
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        const cell = worksheet[cell_ref];
        if (cell && cell.v) {
            localTotalCount++;
            const cleanedNumber = applyExcelFormula(String(cell.v));
            if (cleanedNumber.length > 0) {
                phoneNumbers.push(cleanedNumber);
            } else {
                localEmptyCount++;
            }
        } else {
            localEmptyCount++;
            localTotalCount++;
        }
    }

    return { totalCount: localTotalCount, emptyCount: localEmptyCount };
}

function applyExcelFormula(number) {
    var cleanedNumber = number.replace(/\D/g, ''); // 숫자만 추출
    if (cleanedNumber.startsWith('82010')) {
        cleanedNumber = '8210' + cleanedNumber.slice(4); // '8201'으로 시작하는 번호를 '8210'으로 변경
    } else if (cleanedNumber.startsWith('0')) {
        cleanedNumber = '82' + cleanedNumber.slice(1);
    } else if (!cleanedNumber.startsWith('82')) {
        cleanedNumber = '82' + cleanedNumber;
    }
    return cleanedNumber;
}

function processPhoneNumbers() {
    totalCount = phoneNumbers.length;

    // 1. 빈 값 필터링
    phoneNumbers = phoneNumbers.filter(number => {
        if (number.length > 0) {
            return true;
        } else {
            emptyCount++;
            return false;
        }
    });

    // 2. 전화번호 표준화
    phoneNumbers = phoneNumbers.map(number => {
        number = applyExcelFormula(number);
        return number;
    });

    // 3. 유효한 전화번호 필터링
    phoneNumbers = phoneNumbers.filter(number => {
        if (isValidPhoneNumber(number)) {
            return true;
        } else {
            invalidCount++;
            return false;
        }
    });

    // 중복 제거 단계는 삭제
}

function isValidPhoneNumber(number) {
    return number.length > 10 && number.length <= 12;
}

function displayResults() {
    // 중복 제거
    const uniqueSet = new Set(phoneNumbers);
    uniquePhoneNumbers = Array.from(uniqueSet);
    removedDuplicateCount = phoneNumbers.length - uniquePhoneNumbers.length;

    // 중복된 전화번호의 빈도 계산
    const frequencyMap = phoneNumbers.reduce((acc, number) => {
        acc[number] = (acc[number] || 0) + 1;
        return acc;
    }, {});

    // 빈도가 높은 상위 100개 전화번호 추출
    const topDuplicates = Object.entries(frequencyMap)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 100);

    // 12자리 이상 또는 10자리 이하 번호 유효하지 않은 번호로 처리
    uniquePhoneNumbers = uniquePhoneNumbers.filter(number => {
        if (number.length <= 12 && number.length > 10) {
            return true;
        } else {
            invalidCount++;
            return false;
        }
    });

    if (uniquePhoneNumbers.length > 0) {
        $('#result').hide();
        $('#toggleResultButton').show();

        displayPhoneNumbers();
        displayCounts();
        createCopyButtons();
        createDownloadButtons();
    } else {
        displayNoDataMessage();
    }

    // 상위 100개 중복된 전화번호와 빈도 콘솔에 출력
    console.log('상위 100개 중복된 전화번호:');
    topDuplicates.forEach(([number, count], index) => {
        console.log(`${index + 1}위: ${number} - ${count}회`);
    });
}

function displayPhoneNumbers() {
    var html = '<ul>';
    uniquePhoneNumbers.forEach(number => {
        html += '<li>' + number + '</li>';
    });
    html += '</ul>';

    $('#result').html(html);
}

function displayCounts() {
    $('#count').html('유효한 번호 총 개수: ' + uniquePhoneNumbers.length);
    $('#removedCount').html('중복된 번호 개수: ' + removedDuplicateCount);
    $('#removedCount').append('<br>유효하지 않은 번호 개수: ' + invalidCount);
    $('#removedCount').append('<br>공란의 개수: ' + emptyCount);
    $('#removedCount').append('<br>총 삭제된 개수: ' + (removedDuplicateCount + invalidCount + emptyCount));
    $('#removedCount').append('<br>전체 데이터 개수: ' + totalCount);
}

function createCopyButtons() {
    $('#copyButtons').html('');
    for (let i = 0; i < uniquePhoneNumbers.length; i += 10000) {
        let start = i + 1;
        let end = i + 10000 > uniquePhoneNumbers.length ? uniquePhoneNumbers.length : i + 10000;
        let range = `${start}~${end}`;
        let button = `<button onclick="copyToClipboard(${i}, ${end}, this)">${range}</button>`;
        $('#copyButtons').append(button);
    }
}

function createDownloadButtons() {
    $('#downloadButtons').html('<button id="downloadAllButton">전체 데이터 엑셀 다운로드</button>');
    $('#downloadAllButton').on('click', function () {
        downloadAll(this);
    });

    $('#downloadButtons').append('<button id="toggleDownloadButtons">엑셀 다운로드 접기/펼치기</button>');
    $('#downloadButtons').append('<div id="additionalDownloadButtons" style="display:none;"></div>');

    if (uniquePhoneNumbers.length > 500000) {
        for (let i = 0; i < uniquePhoneNumbers.length; i += 500000) {
            let start = i + 1;
            let end = i + 500000 > uniquePhoneNumbers.length ? uniquePhoneNumbers.length : i + 500000;
            let range = `${start}~${end}`;
            let button = `<button onclick="downloadRange(${i}, ${end}, '${fileName} 수정본 ${start}-${end}.xlsx', this)">${range} 다운로드</button>`;
            $('#additionalDownloadButtons').append(button);
        }
    }

    $('#toggleDownloadButtons').on('click', function () {
        $('#additionalDownloadButtons').toggle();
    });
}

function displayNoDataMessage() {
    $('#result').html('<p>먼저 파일을 업로드해주세요.</p>');
    $('#count').html('');
    $('#removedCount').html('');
    $('#copyButtons').html('');
    $('#downloadButtons').html('');
}

function handleRangeCopy() {
    var startIndex = parseInt($('#startIndex').val()) - 1;
    var endIndex = parseInt($('#endIndex').val());

    if (startIndex >= 0 && endIndex <= uniquePhoneNumbers.length && startIndex < endIndex) {
        copyToClipboard(startIndex, endIndex);
    } else {
        alert('유효한 범위를 입력해주세요.');
    }
}

function copyToClipboard(start, end, button) {
    var textToCopy = uniquePhoneNumbers.slice(start, end).join('\n');
    navigator.clipboard.writeText(textToCopy).then(function () {
        alert('클립보드에 복사되었습니다.');
        if (button) {
            $(button).addClass('clicked');
        }
    }, function (err) {
        console.error('클립보드 복사 실패: ', err);
    });
}

function downloadRange(start, end, filename, button) {
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(uniquePhoneNumbers.slice(start, end).map(number => [number]));
    XLSX.utils.book_append_sheet(wb, ws, 'Unique Phone Numbers');

    XLSX.writeFile(wb, filename);
    if (button) {
        $(button).css('background-color', 'red');
    }
}

function downloadAll(button) {
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(uniquePhoneNumbers.map(number => [number]));
    XLSX.utils.book_append_sheet(wb, ws, 'Unique Phone Numbers');

    var downloadFileName = fileName + ' 전체 수정본.xlsx';
    XLSX.writeFile(wb, downloadFileName);
    if (button) {
        $(button).css('background-color', 'red');
    }
}
