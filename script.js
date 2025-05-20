let excelData = [];
let recognition;

// Biến toàn cục cho việc chọn vùng ảnh
let isSelecting = false;
let isMoving = false;
let isResizing = false;
let startX, startY;
let selectionOverlay = null;
let currentImage = null;
let currentHandle = null;
let originalX, originalY, originalWidth, originalHeight;

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

// Hàm xử lý ảnh và thực hiện OCR
function processImageForOCR(imageDataUrl) {
    console.log('Processing image for OCR...');
    
    const img = new Image();
    img.onload = function() {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        
        // Tính toán kích thước phù hợp
        let width = img.width;
        let height = img.height;
        
        // Giới hạn kích thước tối đa
        const MAX_SIZE = 2048;
        if (width > MAX_SIZE || height > MAX_SIZE) {
            if (width > height) {
                height = Math.round((height * MAX_SIZE) / width);
                width = MAX_SIZE;
            } else {
                width = Math.round((width * MAX_SIZE) / height);
                height = MAX_SIZE;
            }
        }
        
        canvas.width = width;
        canvas.height = height;
        
        // Vẽ ảnh với các tùy chọn để tăng chất lượng
        ctx.drawImage(img, 0, 0, width, height);
        
        // Áp dụng các bộ lọc để cải thiện chất lượng ảnh
        const imageData = ctx.getImageData(0, 0, width, height);
        const data = imageData.data;
        
        // Tăng độ tương phản và làm sắc nét
        for (let i = 0; i < data.length; i += 4) {
            // Chuyển đổi sang grayscale
            const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
            
            // Tăng độ tương phản
            const contrast = 1.5;
            const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
            const newValue = factor * (avg - 128) + 128;
            
            // Áp dụng ngưỡng để làm rõ văn bản
            const threshold = 128;
            const finalValue = newValue > threshold ? 255 : 0;
            
            data[i] = data[i + 1] = data[i + 2] = finalValue;
        }
        
        ctx.putImageData(imageData, 0, 0);
        
        // Chuyển đổi canvas thành Data URL với chất lượng cao
        const processedImageDataUrl = canvas.toDataURL('image/jpeg', 1.0);
        
        // Thực hiện OCR với ảnh đã xử lý
        Tesseract.recognize(
            processedImageDataUrl,
            'vie',
            {
                logger: m => console.log(m),
                tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠàáâãèéêìíòóôõùúăđĩũơƯĂẠẢẤẦẨẪẬẮẰẲẴẶẸẺẼỀỀỂẾưăạảấầẩẫậắằẳẵặẹẻẽềềểếỄỆỈỊỌỎỐỒỔỖỘỚỜỞỠỢỤỦỨỪễệỉịọỏốồổỗộớờởỡợụủứừỬỮỰỲỴÝỶỸửữựỳỵýỷỹ',
                tessedit_pageseg_mode: '6',
                preserve_interword_spaces: '1',
                tessedit_ocr_engine_mode: '1', // Sử dụng LSTM OCR Engine Mode
                tessjs_create_pdf: '0',
                tessjs_create_hocr: '0',
                tessjs_create_tsv: '0',
                tessjs_create_box: '0',
                tessjs_create_unlv: '0',
                tessjs_create_osd: '0'
            }
        ).then(({ data: { text } }) => {
            console.log('Văn bản nhận dạng được:', text);
            // Điền văn bản vào ô tìm kiếm
            document.getElementById('searchInput').value = text.trim();
            showNotification('Đã nhận dạng văn bản từ ảnh');
            
            // Tự động gọi tìm kiếm sau khi nhận dạng thành công
            searchExcelData(text.trim());
            
        }).catch(err => {
            console.error('Lỗi OCR:', err);
            showNotification('Lỗi khi nhận dạng văn bản từ ảnh.', 'error');
        });
    };
    
    img.onerror = function() {
        console.error('Lỗi khi tải ảnh');
        showNotification('Không thể tải ảnh để xử lý.', 'error');
    };
    
    img.src = imageDataUrl;
}

// Xử lý việc chọn file ảnh để OCR
function handleImageSelect(event) {
    const file = event.target.files[0];
    
    if (!file) {
        document.getElementById('imageName').textContent = 'Chưa có ảnh';
        return;
    }
    
    // Cập nhật tên file ảnh được chọn
    document.getElementById('imageName').textContent = file.name;
    showNotification('Đã chọn ảnh. Vui lòng chọn vùng cần xử lý...');
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const imageDataUrl = e.target.result;
        showImageCropModal(imageDataUrl);
    };
    
    reader.onerror = function() {
        console.error('Lỗi khi đọc file ảnh');
        showNotification('Không thể đọc file ảnh. Vui lòng thử lại.', 'error');
    };
    
    reader.readAsDataURL(file);
}

// Hàm xử lý ảnh được chụp trực tiếp
function handleCaptureSelect(event) {
    const file = event.target.files[0];

    if (!file) {
        document.getElementById('captureName').textContent = 'Chưa chụp';
        return;
    }

    document.getElementById('captureName').textContent = 'Đã chụp ảnh';
    showNotification('Đã chụp ảnh. Vui lòng chọn vùng cần xử lý...');

    const reader = new FileReader();

    reader.onload = function(e) {
        const imageDataUrl = e.target.result;
        // Xử lý ảnh chụp trước khi hiển thị modal
        preprocessCapturedImage(imageDataUrl);
    };

    reader.onerror = function() {
        console.error('Lỗi khi đọc ảnh chụp');
        showNotification('Không thể đọc ảnh chụp. Vui lòng thử lại.', 'error');
    };

    reader.readAsDataURL(file);
}

// Hàm tiền xử lý ảnh chụp
function preprocessCapturedImage(imageDataUrl) {
    const img = new Image();
    img.onload = function() {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        
        // Tính toán kích thước phù hợp
        let width = img.width;
        let height = img.height;
        
        // Giới hạn kích thước tối đa
        const MAX_SIZE = 2048;
        if (width > MAX_SIZE || height > MAX_SIZE) {
            if (width > height) {
                height = Math.round((height * MAX_SIZE) / width);
                width = MAX_SIZE;
            } else {
                width = Math.round((width * MAX_SIZE) / height);
                height = MAX_SIZE;
            }
        }
        
        canvas.width = width;
        canvas.height = height;
        
        // Vẽ ảnh với các tùy chọn để tăng chất lượng
        ctx.drawImage(img, 0, 0, width, height);
        
        // Áp dụng các bộ lọc để cải thiện chất lượng ảnh
        const imageData = ctx.getImageData(0, 0, width, height);
        const data = imageData.data;
        
        // Tăng độ tương phản và làm sắc nét
        for (let i = 0; i < data.length; i += 4) {
            // Chuyển đổi sang grayscale
            const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
            
            // Tăng độ tương phản
            const contrast = 1.5;
            const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
            const newValue = factor * (avg - 128) + 128;
            
            // Áp dụng ngưỡng để làm rõ văn bản
            const threshold = 128;
            const finalValue = newValue > threshold ? 255 : 0;
            
            data[i] = data[i + 1] = data[i + 2] = finalValue;
        }
        
        ctx.putImageData(imageData, 0, 0);
        
        // Chuyển đổi canvas thành Data URL với chất lượng cao
        const processedImageDataUrl = canvas.toDataURL('image/jpeg', 1.0);
        
        // Hiển thị modal với ảnh đã xử lý
        showImageCropModal(processedImageDataUrl);
    };
    
    img.onerror = function() {
        console.error('Lỗi khi xử lý ảnh chụp');
        showNotification('Không thể xử lý ảnh chụp. Vui lòng thử lại.', 'error');
    };
    
    img.src = imageDataUrl;
}

// Hàm hiển thị modal chọn vùng ảnh
function showImageCropModal(imageDataUrl) {
    const modal = document.getElementById('imageCropModal');
    const cropImage = document.getElementById('cropImage');
    
    cropImage.src = imageDataUrl;
    currentImage = cropImage;
    modal.style.display = 'block';
    
    // Xóa vùng chọn cũ nếu có
    if (selectionOverlay) {
        selectionOverlay.remove();
        selectionOverlay = null;
    }
    
    // Thêm sự kiện cho việc chọn vùng
    cropImage.onload = function() {
        const container = document.querySelector('.image-container');
        
        // Thêm sự kiện cho cả mouse và touch
        container.addEventListener('mousedown', handleMouseDown);
        container.addEventListener('mousemove', handleMouseMove);
        container.addEventListener('mouseup', handleMouseUp);
        
        // Thêm sự kiện touch
        container.addEventListener('touchstart', handleTouchStart, { passive: false });
        container.addEventListener('touchmove', handleTouchMove, { passive: false });
        container.addEventListener('touchend', handleTouchEnd);
    };
}

// Hàm xử lý sự kiện mousedown
function handleMouseDown(e) {
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    // Kiểm tra nếu click vào handle để resize
    if (e.target.classList.contains('selection-handle')) {
        isResizing = true;
        currentHandle = e.target;
        startX = e.clientX;
        startY = e.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        originalWidth = parseInt(selectionOverlay.style.width);
        originalHeight = parseInt(selectionOverlay.style.height);
        return;
    }
    
    // Kiểm tra nếu click vào vùng chọn để di chuyển
    if (e.target === selectionOverlay) {
        isMoving = true;
        startX = e.clientX;
        startY = e.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        return;
    }
    
    // Tạo vùng chọn mới
    isSelecting = true;
    startX = e.clientX - rect.left;
    startY = e.clientY - rect.top;
    
    // Xóa vùng chọn cũ nếu có
    if (selectionOverlay) {
        selectionOverlay.remove();
    }
    
    // Tạo overlay mới
    selectionOverlay = document.createElement('div');
    selectionOverlay.className = 'selection-overlay';
    selectionOverlay.style.left = startX + 'px';
    selectionOverlay.style.top = startY + 'px';
    
    // Thêm các handle để resize
    const handles = ['nw', 'ne', 'sw', 'se'];
    handles.forEach(pos => {
        const handle = document.createElement('div');
        handle.className = `selection-handle ${pos}`;
        selectionOverlay.appendChild(handle);
    });
    
    container.appendChild(selectionOverlay);
}

// Hàm xử lý sự kiện mousemove
function handleMouseMove(e) {
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    if (isSelecting) {
        const currentX = e.clientX - rect.left;
        const currentY = e.clientY - rect.top;
        
        const width = Math.abs(currentX - startX);
        const height = Math.abs(currentY - startY);
        
        selectionOverlay.style.width = width + 'px';
        selectionOverlay.style.height = height + 'px';
        selectionOverlay.style.left = Math.min(startX, currentX) + 'px';
        selectionOverlay.style.top = Math.min(startY, currentY) + 'px';
    }
    else if (isMoving) {
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;
        
        const newX = originalX + deltaX;
        const newY = originalY + deltaY;
        
        // Giới hạn vùng chọn trong container
        const maxX = rect.width - parseInt(selectionOverlay.style.width);
        const maxY = rect.height - parseInt(selectionOverlay.style.height);
        
        selectionOverlay.style.left = Math.max(0, Math.min(newX, maxX)) + 'px';
        selectionOverlay.style.top = Math.max(0, Math.min(newY, maxY)) + 'px';
    }
    else if (isResizing) {
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;
        
        switch(currentHandle.className) {
            case 'selection-handle nw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle ne':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle sw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                break;
            case 'selection-handle se':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                break;
        }
        
        // Đảm bảo kích thước tối thiểu
        const minSize = 20;
        if (parseInt(selectionOverlay.style.width) < minSize) {
            selectionOverlay.style.width = minSize + 'px';
        }
        if (parseInt(selectionOverlay.style.height) < minSize) {
            selectionOverlay.style.height = minSize + 'px';
        }
    }
}

// Hàm xử lý sự kiện mouseup
function handleMouseUp() {
    isSelecting = false;
    isMoving = false;
    isResizing = false;
    currentHandle = null;
}

// Hàm xử lý sự kiện touchstart
function handleTouchStart(e) {
    e.preventDefault(); // Ngăn chặn scroll khi touch
    const touch = e.touches[0];
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    // Kiểm tra nếu touch vào handle để resize
    if (e.target.classList.contains('selection-handle')) {
        isResizing = true;
        currentHandle = e.target;
        startX = touch.clientX;
        startY = touch.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        originalWidth = parseInt(selectionOverlay.style.width);
        originalHeight = parseInt(selectionOverlay.style.height);
        return;
    }
    
    // Kiểm tra nếu touch vào vùng chọn để di chuyển
    if (e.target === selectionOverlay) {
        isMoving = true;
        startX = touch.clientX;
        startY = touch.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        return;
    }
    
    // Tạo vùng chọn mới
    isSelecting = true;
    startX = touch.clientX - rect.left;
    startY = touch.clientY - rect.top;
    
    // Xóa vùng chọn cũ nếu có
    if (selectionOverlay) {
        selectionOverlay.remove();
    }
    
    // Tạo overlay mới
    selectionOverlay = document.createElement('div');
    selectionOverlay.className = 'selection-overlay';
    selectionOverlay.style.left = startX + 'px';
    selectionOverlay.style.top = startY + 'px';
    
    // Thêm các handle để resize
    const handles = ['nw', 'ne', 'sw', 'se'];
    handles.forEach(pos => {
        const handle = document.createElement('div');
        handle.className = `selection-handle ${pos}`;
        selectionOverlay.appendChild(handle);
    });
    
    container.appendChild(selectionOverlay);
}

// Hàm xử lý sự kiện touchmove
function handleTouchMove(e) {
    e.preventDefault(); // Ngăn chặn scroll khi touch
    const touch = e.touches[0];
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    if (isSelecting) {
        const currentX = touch.clientX - rect.left;
        const currentY = touch.clientY - rect.top;
        
        const width = Math.abs(currentX - startX);
        const height = Math.abs(currentY - startY);
        
        selectionOverlay.style.width = width + 'px';
        selectionOverlay.style.height = height + 'px';
        selectionOverlay.style.left = Math.min(startX, currentX) + 'px';
        selectionOverlay.style.top = Math.min(startY, currentY) + 'px';
    }
    else if (isMoving) {
        const deltaX = touch.clientX - startX;
        const deltaY = touch.clientY - startY;
        
        const newX = originalX + deltaX;
        const newY = originalY + deltaY;
        
        // Giới hạn vùng chọn trong container
        const maxX = rect.width - parseInt(selectionOverlay.style.width);
        const maxY = rect.height - parseInt(selectionOverlay.style.height);
        
        selectionOverlay.style.left = Math.max(0, Math.min(newX, maxX)) + 'px';
        selectionOverlay.style.top = Math.max(0, Math.min(newY, maxY)) + 'px';
    }
    else if (isResizing) {
        const deltaX = touch.clientX - startX;
        const deltaY = touch.clientY - startY;
        
        switch(currentHandle.className) {
            case 'selection-handle nw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle ne':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle sw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                break;
            case 'selection-handle se':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                break;
        }
        
        // Đảm bảo kích thước tối thiểu
        const minSize = 20;
        if (parseInt(selectionOverlay.style.width) < minSize) {
            selectionOverlay.style.width = minSize + 'px';
        }
        if (parseInt(selectionOverlay.style.height) < minSize) {
            selectionOverlay.style.height = minSize + 'px';
        }
    }
}

// Hàm xử lý sự kiện touchend
function handleTouchEnd() {
    isSelecting = false;
    isMoving = false;
    isResizing = false;
    currentHandle = null;
}

// Hàm cắt ảnh theo vùng đã chọn
function cropImage() {
    if (!selectionOverlay || !currentImage) return;
    
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    
    const rect = selectionOverlay.getBoundingClientRect();
    const imageRect = currentImage.getBoundingClientRect();
    
    // Tính toán tỷ lệ giữa ảnh gốc và ảnh hiển thị
    const scaleX = currentImage.naturalWidth / imageRect.width;
    const scaleY = currentImage.naturalHeight / imageRect.height;
    
    // Tính toán vị trí và kích thước thực tế của vùng cắt
    const cropX = (rect.left - imageRect.left) * scaleX;
    const cropY = (rect.top - imageRect.top) * scaleY;
    const cropWidth = rect.width * scaleX;
    const cropHeight = rect.height * scaleY;
    
    canvas.width = cropWidth;
    canvas.height = cropHeight;
    
    // Vẽ phần ảnh đã cắt
    ctx.drawImage(
        currentImage,
        cropX, cropY, cropWidth, cropHeight,
        0, 0, cropWidth, cropHeight
    );
    
    // Chuyển đổi canvas thành Data URL
    const croppedImageDataUrl = canvas.toDataURL('image/jpeg');
    
    // Đóng modal
    document.getElementById('imageCropModal').style.display = 'none';
    
    // Xử lý ảnh đã cắt
    processImageForOCR(croppedImageDataUrl);
}

window.onload = () => {
    // Đăng ký sự kiện cho nút chọn file Excel
    document.getElementById('fileInput').addEventListener('change', handleFileSelect);
    
    // Đăng ký sự kiện cho nút chọn file ảnh
    document.getElementById('imageInput').addEventListener('change', handleImageSelect);
    
    // Thêm trình lắng nghe sự kiện paste
    document.body.addEventListener('paste', handlePaste);
    
    // Đăng ký sự kiện cho nút chụp ảnh
    document.getElementById('captureInput').addEventListener('change', handleCaptureSelect);
    
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
    
    // Thêm event listeners cho modal
    document.querySelector('.close').onclick = function() {
        document.getElementById('imageCropModal').style.display = 'none';
    };
    
    document.getElementById('cropButton').onclick = cropImage;
    
    document.getElementById('cancelCropButton').onclick = function() {
        document.getElementById('imageCropModal').style.display = 'none';
    };
    
    // Đóng modal khi click bên ngoài
    window.onclick = function(event) {
        const modal = document.getElementById('imageCropModal');
        if (event.target == modal) {
            modal.style.display = 'none';
        }
    };
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