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

    console.log(`파일 업로드 시작: ${files.length}개의 파일`);

    // 모든 파일을 병렬로 읽기
    await Promise.all(files.map(file => readFile(file)));

    processPhoneNumbers();
    enableConfirmButton();
    $('#confirmButton').trigger('click');
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
                handleWorkbook(workbook);
                console.log(`파일 처리 완료: ${file.name}`);
                resolve();
            } catch (error) {
                console.error('파일을 읽는 중 오류가 발생했습니다.', error);
                alert('파일을 읽는 중 오류가 발생했습니다.');
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function handleWorkbook(workbook) {
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet['!ref']) return;
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const columnData = getColumnDataWithMaxMatches(worksheet, range);
        if (columnData.targetColumn >= 0) {
            extractPhoneNumbersFromColumn(worksheet, range, columnData.targetColumn);
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
        /10\d{7}/
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
    for (let R = range.s.r; R <= range.e.r; ++R) {
        const cell_address = { c: targetColumn, r: R };
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        const cell = worksheet[cell_ref];
        if (cell && cell.v) {
            const cleanedNumber = applyExcelFormula(String(cell.v));
            if (cleanedNumber.length > 0) {
                phoneNumbers.push(cleanedNumber);
            } else {
                emptyCount++;
            }
        } else {
            emptyCount++;
        }
    }
}

function applyExcelFormula(number) {
    // 엑셀 수식을 JavaScript로 변환
    var digits = number.replace(/\D/g, ''); // 숫자만 추출
    var result = 0;
    for (let i = 0; i < digits.length; i++) {
        result += digits[digits.length - 1 - i] * Math.pow(10, i);
    }
    return '82' + result;
}

function processPhoneNumbers() {
    phoneNumbers = phoneNumbers.filter(number => {
        if (number.length > 0) {
            return true;
        } else {
            emptyCount++;
            return false;
        }
    });

    phoneNumbers = phoneNumbers.map(number => {
        number = number.replace(/^0+/, '');
        if (!number.startsWith('82')) {
            number = '82' + number;
        }
        return number;
    });

    const validPhoneNumbers = phoneNumbers.filter(number => {
        if (number.length === 12 || (number.startsWith('82') && (number.length === 12 || number.length === 13))) {
            return true;
        } else {
            invalidCount++;
            return false;
        }
    });

    const uniqueSet = new Set(validPhoneNumbers);
    uniquePhoneNumbers = Array.from(uniqueSet);
    removedDuplicateCount = validPhoneNumbers.length - uniquePhoneNumbers.length;

    // 전체 데이터 개수 계산
    totalCount = uniquePhoneNumbers.length + removedDuplicateCount + invalidCount;
}

function displayResults() {
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
