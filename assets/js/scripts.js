var phoneNumbers = [];
var uniquePhoneNumbers = [];
var invalidCount = 0;
var totalCount = 0;
var removedDuplicateCount = 0;
var fileName = '';

$(document).ready(function () {
    $('#fileUpload').on('change', function (e) {
        $('#confirmButton').prop('disabled', true); // 확인 버튼 비활성화
        $('#status').html(''); // 상태 초기화
        phoneNumbers = [];
        uniquePhoneNumbers = [];
        invalidCount = 0;
        totalCount = 0;
        removedDuplicateCount = 0;

        var files = e.target.files;
        fileName = files[0].name.split('.').slice(0, -1).join('.'); // 확장자 제외한 파일 이름 저장
        
        var fileCount = files.length;
        var processedFiles = 0;

        Array.from(files).forEach(file => {
            var reader = new FileReader();

            reader.onload = function (e) {
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, { type: 'array' });

                workbook.SheetNames.forEach(sheetName => {
                    var worksheet = workbook.Sheets[sheetName];
                    if (!worksheet['!ref']) {
                        updateStatus(`빈 시트이거나 참조가 없습니다: ${sheetName}`);
                        return;
                    }
                    var range = XLSX.utils.decode_range(worksheet['!ref']);
                    var regexPatterns = [
                        /010[- ]?\d{4}[- ]?\d{4}/,   // 010-1234-5678 or 010 1234 5678
                        /010\d{7,8}/,                // 01012345678 or 010123456789
                        /82\d{2}[- ]?\d{3,4}[- ]?\d{4}/,  // 8210-1234-5678 or 8210 1234 5678
                        /82\d{2}\d{7,8}/,            // 821012345678 or 8210123456789
                        /01[1-9][- ]?\d{3,4}[- ]?\d{4}/,   // 011-123-4567 or 011 123 4567
                        /10[- ]?\d{4}[- ]?\d{4}/,    // 10-1234-5678 or 10 1234 5678
                        /10\d{7}/                    // 1012345678
                    ];

                    let maxMatchCount = 0;
                    let targetColumn = -1;

                    // 모든 열을 검사하여 가장 많은 휴대폰 번호를 포함한 열 찾기
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        let columnMatchCount = 0;
                        let matchedNumbers = new Set();

                        for (let R = range.s.r; R <= range.e.r; ++R) {
                            const cell_address = { c: C, r: R };
                            const cell_ref = XLSX.utils.encode_cell(cell_address);
                            const cell = worksheet[cell_ref];
                            if (cell && cell.v) {
                                const cellValue = String(cell.v); // 문자열로 변환
                                regexPatterns.forEach(pattern => {
                                    const matches = cellValue.match(pattern);
                                    if (matches) {
                                        matches.forEach(match => matchedNumbers.add(match));
                                    }
                                });
                            }
                        }

                        columnMatchCount = matchedNumbers.size;
                        updateStatus(`열 ${XLSX.utils.encode_col(C)}에서 매칭된 휴대폰 번호 수: ${columnMatchCount}`);

                        if (columnMatchCount > maxMatchCount) {
                            maxMatchCount = columnMatchCount;
                            targetColumn = C;
                        }
                    }

                    if (targetColumn >= 0) {
                        updateStatus(`가장 많은 휴대폰 번호가 있는 열: ${XLSX.utils.encode_col(targetColumn)} (매칭된 번호 수: ${maxMatchCount})`);
                        for (let R = range.s.r; R <= range.e.r; ++R) {
                            const cell_address = { c: targetColumn, r: R };
                            const cell_ref = XLSX.utils.encode_cell(cell_address);
                            const cell = worksheet[cell_ref];
                            if (cell && cell.v) {
                                const cleanedNumber = String(cell.v).replace(/[^0-9]/g, '');
                                // console.log(`Original: ${cell.v}, Cleaned: ${cleanedNumber}`); // 디버그 로그 추가
                                const lastEightDigits = cleanedNumber.slice(-8);
                                if (lastEightDigits.length === 8) {
                                    phoneNumbers.push('8210' + lastEightDigits);
                                } else {
                                    invalidCount++;
                                }
                            }
                        }
                    } else {
                        updateStatus('휴대폰 번호를 찾을 수 없습니다.');
                    }
                });

                // 파일 처리 완료 체크
                processedFiles++;
                if (processedFiles === fileCount) {
                    // 중복 제거
                    var uniqueSet = new Set(phoneNumbers);
                    uniquePhoneNumbers = Array.from(uniqueSet);
                    removedDuplicateCount = phoneNumbers.length - uniquePhoneNumbers.length;
                    updateStatus(`중복된 번호 개수: ${removedDuplicateCount}, 유효하지 않은 번호 개수: ${invalidCount}`);
                    
                    $('#confirmButton').prop('disabled', false); // 확인 버튼 활성화
                }
            };

            reader.onerror = function () {
                alert('파일을 읽는 중 오류가 발생했습니다.');
            };

            reader.readAsArrayBuffer(file);
        });
    });

    $('#confirmButton').on('click', function () {
        if (uniquePhoneNumbers.length > 0) {
            $('#result').hide();
            $('#toggleResultButton').show();
            $('#downloadButton').show();

            var html = '<ul>';
            uniquePhoneNumbers.forEach(number => {
                html += '<li>' + number + '</li>';
            });
            html += '</ul>';

            $('#result').html(html);
            $('#count').html('유효한 번호 총 개수: ' + uniquePhoneNumbers.length);
            $('#removedCount').html('중복된 번호 개수: ' + removedDuplicateCount);
            $('#removedCount').append('<br>유효하지 않은 번호 개수: ' + invalidCount);
            $('#removedCount').append('<br>총 삭제된 개수: ' + (removedDuplicateCount + invalidCount));
            $('#removedCount').append('<br>전체 데이터 개수: ' + (uniquePhoneNumbers.length + removedDuplicateCount + invalidCount));

            $('#copyButtons').html('');
            for (let i = 0; i < uniquePhoneNumbers.length; i += 10000) {
                let start = i + 1;
                let end = i + 10000 > uniquePhoneNumbers.length ? uniquePhoneNumbers.length : i + 10000;
                let range = `${start}~${end}`;
                let button = `<button onclick="copyToClipboard(${i}, ${end}, this)">${range}</button>`;
                $('#copyButtons').append(button);
            }
        } else {
            $('#result').html('<p>먼저 파일을 업로드해주세요.</p>');
            $('#count').html('');
            $('#removedCount').html('');
            $('#copyButtons').html('');
        }
    });

    $('#toggleResultButton').on('click', function () {
        $('#result').toggle();
    });

    $('#downloadButton').on('click', function () {
        var wb = XLSX.utils.book_new();
        var ws = XLSX.utils.aoa_to_sheet(uniquePhoneNumbers.map(number => [number]));
        XLSX.utils.book_append_sheet(wb, ws, 'Unique Phone Numbers');

        var downloadFileName = fileName + ' 수정본.xlsx';
        XLSX.writeFile(wb, downloadFileName);
    });

    $('#rangeCopyButton').on('click', function () {
        var startIndex = parseInt($('#startIndex').val()) - 1;
        var endIndex = parseInt($('#endIndex').val());

        if (startIndex >= 0 && endIndex <= uniquePhoneNumbers.length && startIndex < endIndex) {
            copyToClipboard(startIndex, endIndex);
        } else {
            alert('유효한 범위를 입력해주세요.');
        }
    });
});

function updateStatus(message) {
    $('#status').html(`<p>${message}</p>`);
    setTimeout(() => {
        $('#status').html('');
    }, 100); // 메시지를 잠시 후에 지우기
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
