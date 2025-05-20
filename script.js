let excelData = [];
let recognition;

// Khởi tạo đối tượng nhận diện giọng nói
function initSpeechRecognition() {
    // Tạo đối tượng nhận diện giọng nói nếu hỗ trợ
    if ('SpeechRecognition' in window || 'webkitSpeechRecognition' in window) {
        recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
        recognition.lang = 'vi-VN';
        recognition.interimResults = false;
        recognition.maxAlternatives = 1;
        
        // Xử lý kết quả nhận diện
        recognition.onresult = (event) => {
            const transcript = event.results[0][0].transcript.toLowerCase();
            const cleanedTranscript = transcript.replace(/[.,!?;:]/g, '').trim();
            console.log('Văn bản đã làm sạch:', cleanedTranscript);
            console.log('Nhận diện văn bản:', transcript);
            
            // Kiểm tra nếu là lệnh xóa
            if (transcript.includes('xóa') || transcript.includes('xoá') || 
                transcript.includes('làm mới') || transcript.includes('xóa kết quả')) {
                clearResults();
                showNotification('Đã xóa kết quả theo lệnh giọng nói');
            } 
            // Kiểm tra nếu là lệnh hiển thị tất cả (phần trăm phần trăm)
            else if (transcript.includes('phần trăm phần trăm') || 
                     transcript.includes('hiển thị tất cả') || 
                     transcript.includes('hiện tất cả') ||
                     transcript.includes('tất cả') ||
                     transcript.includes('hiện hết') ||
                     transcript.includes('xem hết')) {
                document.getElementById('searchInput').value = '%%';
                searchExcelData('%%');
                showNotification('Hiển thị tất cả dữ liệu theo yêu cầu');
            } else {
                // Nếu không phải lệnh đặc biệt, thì xử lý như tìm kiếm bình thường
                document.getElementById('searchInput').value = cleanedTranscript;
                searchExcelData(cleanedTranscript);
            }
            
            // Khôi phục nút sau khi có kết quả
            resetListeningButton();
        };

        recognition.onerror = (event) => {
            console.error('Lỗi trong quá trình nhận diện:', event.error);
            resetListeningButton();
        };
        
        recognition.onend = () => {
            resetListeningButton();
        };
    } else {
        showNotification('Trình duyệt của bạn không hỗ trợ nhận diện giọng nói', 'error');
        document.getElementById('startListeningButton').disabled = true;
    }
}

// Hàm để khởi động nhận diện giọng nói
function startListening() {
    if (!recognition) {
        initSpeechRecognition();
    }
    
    try {
        recognition.start();
        // Thêm thông báo đang lắng nghe
        document.getElementById('startListeningButton').textContent = 'Đang lắng nghe...';
        document.getElementById('startListeningButton').style.backgroundColor = '#4a9b4a';
        
        // Hiển thị hướng dẫn lệnh giọng nói
        showVoiceCommands();
    } catch (error) {
        console.error('Lỗi khi bắt đầu nhận diện:', error);
        resetListeningButton();
    }
}

// Khôi phục trạng thái nút lắng nghe
function resetListeningButton() {
    document.getElementById('startListeningButton').textContent = 'Bắt Đầu Nói';
    document.getElementById('startListeningButton').style.backgroundColor = '#8A1538';
    // Ẩn hướng dẫn lệnh giọng nói sau khi hoàn thành
    hideVoiceCommands();
}

// Hiển thị hướng dẫn lệnh giọng nói
function showVoiceCommands() {
    let voiceCommandsHelp = document.getElementById('voiceCommandsHelp');
    
    if (!voiceCommandsHelp) {
        voiceCommandsHelp = document.createElement('div');
        voiceCommandsHelp.id = 'voiceCommandsHelp';
        voiceCommandsHelp.innerHTML = `
            <div style="font-size: 13px; color: #666; margin-top: 5px; padding: 5px 10px; background-color: #f0f2f5; border-radius: 15px; display: inline-block;">
                Nói <strong>"xóa"</strong> hoặc <strong>"xóa kết quả"</strong> để làm mới kết quả<br>
                Nói <strong>"tất cả"</strong> hoặc <strong>"hiện tất cả"</strong> để hiển thị toàn bộ dữ liệu
            </div>
        `;
        
        // Thêm sau nút lắng nghe
        const controlsContainer = document.getElementById('controls-container');
        controlsContainer.appendChild(voiceCommandsHelp);
    } else {
        voiceCommandsHelp.style.display = 'block';
    }
}

// Ẩn hướng dẫn lệnh giọng nói
function hideVoiceCommands() {
    const voiceCommandsHelp = document.getElementById('voiceCommandsHelp');
    if (voiceCommandsHelp) {
        voiceCommandsHelp.style.display = 'none';
    }
}

// Xóa kết quả tìm kiếm và làm trống ô tìm kiếm
function clearResults() {
    document.getElementById('searchInput').value = '';
    document.getElementById('resultArea').innerHTML = '';
}

// Hàm để tìm kiếm dữ liệu trong file Excel
function searchExcelData(searchTerm) {
    if (excelData.length === 0) {
        showNotification('Vui lòng chọn file Excel trước khi tìm kiếm.', 'warning');
        return;
    }
    
    if (!searchTerm || searchTerm.trim() === '') {
        showNotification('Vui lòng nhập từ khóa tìm kiếm.', 'info');
        return;
    }

    let results;
    const isShowAll = searchTerm.trim() === '%%';
    
    // Nếu searchTerm là "%%", hiển thị tất cả dữ liệu
    if (isShowAll) {
        results = excelData.filter(row => row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== ''));
    } else {
        // Tìm kiếm thông thường
        results = excelData.filter(row => 
            row.some(cell => cell && cell.toString().toLowerCase().includes(searchTerm.toLowerCase()))
        );
    }

    const resultArea = document.getElementById('resultArea');
    resultArea.innerHTML = ''; // Xóa kết quả cũ

    if (results.length > 0) {
        // Hiển thị số lượng kết quả tìm thấy
        const resultCount = document.createElement('div');
        resultCount.className = 'result-count';
        
        if (isShowAll) {
            resultCount.textContent = `Hiển thị tất cả ${results.length} bản ghi dữ liệu`;
        } else {
            resultCount.textContent = `Tìm thấy ${results.length} kết quả cho "${searchTerm}"`;
        }
        
        resultCount.style.marginBottom = '15px';
        resultCount.style.fontWeight = 'bold';
        resultCount.style.color = '#8A1538';
        resultCount.style.borderBottom = '1px solid #eaeaea';
        resultCount.style.paddingBottom = '8px';
        resultArea.appendChild(resultCount);
        
        // Hiển thị các kết quả tìm kiếm
        results.forEach((row, index) => {
            const resultRow = document.createElement('div');
            
            // Format dữ liệu kết quả để dễ đọc hơn
            const formattedData = formatResultRow(row, isShowAll ? "" : searchTerm);
            resultRow.innerHTML = formattedData;
            
            // Thêm số thứ tự cho mỗi kết quả
            resultRow.dataset.index = index + 1;
            
            resultArea.appendChild(resultRow);
        });
    } else {
        const noResult = document.createElement('div');
        noResult.className = 'no-result';
        noResult.innerHTML = `<p style="text-align: center; color: #888;">Không tìm thấy kết quả cho "${searchTerm}"</p>`;
        resultArea.appendChild(noResult);
    }
}

// Hàm định dạng kết quả để dễ đọc hơn và chỉ hiển thị cột F
function formatResultRow(row, searchTerm) {
    let result = '';
    
    // Chỉ lấy và hiển thị giá trị từ cột F (index 5)
    const cellF = row[5];
    
    if (cellF !== null && cellF !== undefined && cellF.toString().trim() !== '') {
        const cellStr = cellF.toString();
        const lowerCellStr = cellStr.toLowerCase();
        const lowerSearchTerm = searchTerm.toLowerCase();
        
        let formattedCell = cellStr;
        
        // Highlight từ khóa tìm kiếm nếu có
        if (lowerSearchTerm && lowerSearchTerm !== '%%' && lowerCellStr.includes(lowerSearchTerm)) {
            const startIndex = lowerCellStr.indexOf(lowerSearchTerm);
            const endIndex = startIndex + searchTerm.length;
            
            formattedCell = cellStr.substring(0, startIndex) + 
                   `<span style="background-color: #ffe2e8; font-weight: bold;">${cellStr.substring(startIndex, endIndex)}</span>` + 
                   cellStr.substring(endIndex);
        }
        
        result = `<span class="column-f-highlight" style="color: white; font-weight: bold; font-size: 1.1em; background-color: #8A1538; padding: 4px 10px; border-radius: 4px; box-shadow: 0 2px 5px rgba(138, 21, 56, 0.4); margin: 0 3px; display: inline-block; position: relative; max-width: 100%; overflow-wrap: break-word;">
                 <span class="column-f-label" style="position: absolute; top: -10px; left: 0; background-color: #FFD700; color: #8A1538; font-size: 9px; padding: 1px 4px; border-radius: 3px; font-weight: bold;">CỘT F</span>
                 ${formattedCell}
               </span>`;
    }
    
    return `<div style="margin: 10px 0;">${result}</div>`;
}

// Hiển thị thông báo
function showNotification(message, type = 'info') {
    // Tạo thông báo
    const notification = document.createElement('div');
    notification.className = 'notification ' + type;
    notification.textContent = message;
    
    // Style cho thông báo
    notification.style.position = 'fixed';
    notification.style.top = '20px';
    notification.style.left = '50%';
    notification.style.transform = 'translateX(-50%)';
    notification.style.padding = '10px 20px';
    notification.style.borderRadius = '5px';
    notification.style.zIndex = '1000';
    notification.style.boxShadow = '0 3px 6px rgba(0,0,0,0.2)';
    
    // Màu sắc dựa trên loại thông báo
    if (type === 'warning') {
        notification.style.backgroundColor = '#fff3cd';
        notification.style.color = '#856404';
        notification.style.border = '1px solid #ffeeba';
    } else if (type === 'error') {
        notification.style.backgroundColor = '#f8d7da';
        notification.style.color = '#721c24';
        notification.style.border = '1px solid #f5c6cb';
    } else {
        notification.style.backgroundColor = '#d4edda';
        notification.style.color = '#155724';
        notification.style.border = '1px solid #c3e6cb';
    }
    
    // Thêm vào body
    document.body.appendChild(notification);
    
    // Xóa thông báo sau 3 giây
    setTimeout(() => {
        notification.style.opacity = '0';
        notification.style.transition = 'opacity 0.5s';
        setTimeout(() => document.body.removeChild(notification), 500);
    }, 3000);
}

// Tải dữ liệu từ file Excel khi chọn file
function handleFileSelect(event) {
    const file = event.target.files[0];
    
    if (!file) return;
    
    // Cập nhật tên file được chọn
    document.getElementById('fileName').textContent = file.name;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            console.log('Đã tải dữ liệu Excel:', excelData.length, 'dòng');
            showNotification(`Đã tải thành công file ${file.name} với ${excelData.length} dòng dữ liệu.`);
        } catch (error) {
            console.error('Lỗi khi xử lý file:', error);
            showNotification('Không thể xử lý file Excel. Vui lòng kiểm tra định dạng file.', 'error');
        }
    };
    
    reader.onerror = function() {
        console.error('Lỗi khi đọc file');
        showNotification('Không thể đọc file. Vui lòng thử lại.', 'error');
    };
    
    reader.readAsArrayBuffer(file);
}

// Tải file Excel mặc định khi trang web được tải
function loadDefaultExcelFile() {
    try {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', 'data.xlsx', true);
        xhr.responseType = 'arraybuffer';

        xhr.onload = function () {
            if (xhr.status === 200) {
                const data = new Uint8Array(xhr.response);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                console.log('Đã tải dữ liệu Excel mặc định:', excelData.length, 'dòng');
                document.getElementById('fileName').textContent = 'data.xlsx';
                showNotification(`Đã tải thành công file mặc định với ${excelData.length} dòng dữ liệu.`);
            } else {
                console.warn('Không thể tải file Excel mặc định:', xhr.status);
                showNotification('Không tìm thấy file mặc định. Vui lòng chọn file Excel.', 'warning');
            }
        };

        xhr.onerror = function() {
            console.warn('Lỗi khi tải file Excel mặc định');
            showNotification('Không tìm thấy file mặc định. Vui lòng chọn file Excel.', 'warning');
        };

        xhr.send();
    } catch (error) {
        console.warn('Lỗi khi tải file Excel mặc định:', error);
    }
}

window.onload = () => {
    // Đăng ký sự kiện cho nút chọn file
    document.getElementById('fileInput').addEventListener('change', handleFileSelect);
    
    // Tải file Excel mặc định
    loadDefaultExcelFile();
    
    // Khởi tạo nhận diện giọng nói
    initSpeechRecognition();
    
    // Đăng ký sự kiện cho các nút khác
    document.getElementById('startListeningButton').onclick = startListening;
    
    document.getElementById('clearResultsButton').onclick = clearResults;
    
    // Gọi hàm tìm kiếm khi nhấn nút tìm kiếm
    document.getElementById('searchButton').onclick = () => {
        const searchTerm = document.getElementById('searchInput').value;
        searchExcelData(searchTerm); // Gọi hàm tìm kiếm
    };
    
    // Tìm kiếm khi nhấn phím Enter
    document.getElementById('searchInput').addEventListener('keypress', function(event) {
        if (event.key === 'Enter') {
            event.preventDefault(); // Ngăn chặn hành động mặc định
            const searchTerm = document.getElementById('searchInput').value;
            searchExcelData(searchTerm); // Gọi hàm tìm kiếm
        }
    });
};

function displayResults(data) {
    const resultArea = document.getElementById("resultArea");
    resultArea.innerHTML = "";

    data.forEach((row, rowIndex) => {
        if (rowIndex === 0) return; // Bỏ qua header

        const div = document.createElement("div");
        let rowContent = "";

        row.forEach((cell, colIndex) => {
            // Nếu là cột F (index 5), bọc trong span với class nổi bật
            if (colIndex === 5) {
                rowContent += `<span class="column-f-highlight"><span class="column-f-label">Cột F</span>${cell}</span> `;
            } else {
                rowContent += `<span>${cell}</span> `;
            }
        });

        div.innerHTML = rowContent.trim();
        resultArea.appendChild(div);
    });
}