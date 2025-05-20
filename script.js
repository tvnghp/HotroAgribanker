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
async function processImageForOCR(imageDataUrl, isPreview = false) {
    console.log(`processImageForOCR called (isPreview: ${isPreview})`);
    if (!isPreview) {
        showNotification('Đang nhận dạng văn bản từ ảnh...');
    } else {
         console.log('Preprocessing image for preview...');
    }

    return new Promise((resolve, reject) => {
        const img = new Image();
        
        img.onload = function() {
            console.log('Original image loaded for processing/preview');
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
            
            // Vẽ ảnh gốc lên canvas
            ctx.drawImage(img, 0, 0, width, height);
            
            // Lấy dữ liệu pixel
            let imageData = ctx.getImageData(0, 0, width, height);
            let data = imageData.data;
            
            // --- Áp dụng tiền xử lý ảnh ---

            // 1. Chuyển đổi sang grayscale
            for (let i = 0; i < data.length; i += 4) {
                const gray = 0.2126 * data[i] + 0.7152 * data[i + 1] + 0.0722 * data[i + 2];
                data[i] = data[i + 1] = data[i + 2] = gray;
            }

            // Cập nhật dữ liệu grayscale lên canvas tạm thời để áp dụng bộ lọc
            ctx.putImageData(imageData, 0, 0);
            const grayImageData = ctx.getImageData(0, 0, width, height);
            const grayData = grayImageData.data;
            
            // 2. Áp dụng làm mờ Gaussian nhẹ để giảm nhiễu
            const blurRadius = 1; // Điều chỉnh bán kính làm mờ nếu cần
            const blurredData = new Uint8ClampedArray(grayData.length);

            for (let y = 0; y < height; y++) {
                for (let x = 0; x < width; x++) {
                    const i = (y * width + x) * 4;
                    let sum = 0;
                    let count = 0;

                    for (let wy = Math.max(0, y - blurRadius); wy < Math.min(height, y + blurRadius + 1); wy++) {
                        for (let wx = Math.max(0, x - blurRadius); wx < Math.min(width, x + blurRadius + 1); wx++) {
                            const wi = (wy * width + wx) * 4;
                            sum += grayData[wi];
                            count++;
                        }
                    }
                    blurredData[i] = blurredData[i+1] = blurredData[i+2] = sum / count;
                    blurredData[i+3] = 255; // Alpha channel
                }
            }
            
            // Cập nhật dữ liệu đã làm mờ lên canvas
            ctx.putImageData(new ImageData(blurredData, width, height), 0, 0);
            const blurredImageData = ctx.getImageData(0, 0, width, height);
            const blurredDataForThresholding = blurredImageData.data;

            // 3. Áp dụng Thresholding thích ứng
            const windowSize = 15; // Kích thước cửa sổ cho ngưỡng thích ứng (có thể tăng nhẹ)
            const c = 2; // Hằng số trừ đi từ ngưỡng (có thể giảm nhẹ)

            for (let y = 0; y < height; y++) {
                for (let x = 0; x < width; x++) {
                    const i = (y * width + x) * 4;
                    let sum = 0;
                    let count = 0;

                    // Tính tổng giá trị xám trong cửa sổ từ ảnh đã làm mờ
                    for (let wy = Math.max(0, y - windowSize); wy < Math.min(height, y + windowSize + 1); wy++) {
                        for (let wx = Math.max(0, x - windowSize); wx < Math.min(width, x + windowSize + 1); wx++) {
                            const wi = (wy * width + wx) * 4;
                            sum += blurredDataForThresholding[wi];
                            count++;
                        }
                    }

                    // Tính ngưỡng thích ứng và áp dụng
                    const threshold = (sum / count) - c;
                    const grayValue = blurredDataForThresholding[i];

                    if (grayValue > threshold) {
                        data[i] = data[i + 1] = data[i + 2] = 255; // White
                    } else {
                        data[i] = data[i + 1] = data[i + 2] = 0; // Black
                    }
                }
            }

            // Cập nhật lại dữ liệu pixel nhị phân lên canvas
            ctx.putImageData(imageData, 0, 0);
            
            // --- Kết thúc tiền xử lý ---
            
            // Lấy Data URL của ảnh đã xử lý
            const processedImageDataUrl = canvas.toDataURL('image/jpeg', 1.0);
            console.log('Image preprocessing complete. Processed Data URL obtained', processedImageDataUrl.substring(0, 50) + '...');

            if (isPreview) {
                // Nếu chỉ xem trước, trả về Data URL của ảnh đã xử lý
                console.log('Returning processed image data URL for preview');
                resolve(processedImageDataUrl);
            } else {
                // Nếu thực hiện OCR, gọi Tesseract.recognize với ảnh đã xử lý
                console.log('Starting Tesseract recognition...');
                Tesseract.recognize(
                    processedImageDataUrl,
                    'vie',
                    {
                        logger: m => console.log(m),
                        tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠàáâãèéêìòóôõùúăđĩũơƯĂẠẢẤẦẨẪẬẮẰẲẴẶẸẺẼỀỀỂẾưăạảấầẩẫậắằẳẵặẹẻẽềềểếỄỆỈỊỌỎỐỒỔỖỘỚỜỞỠỢỤỦỨỪễệỉịọỏốồổỗộớờởỡợụủứừỬỮỰỲỴÝỶỸửữựỳỵýỷỹ',
                        tessedit_pageseg_mode: '6', // Page segmentation mode: Assume a single uniform block of text.
                        preserve_interword_spaces: '1',
                        tessedit_ocr_engine_mode: '1', // Sử dụng LSTM OCR Engine Mode
                        tessjs_create_pdf: '0',
                        tessjs_create_hocr: '0',
                        tessjs_create_tsv: '0',
                        tessjs_create_box: '0',
                        tessjs_create_unlv: '0',
                        tessjs_create_osd: '0'
                    }
                ).then(({ data }) => {
                    console.log('Tesseract recognition complete. Data:', data);
                    
                    let recognizedText = '';
                    if (data && data.text) {
                        recognizedText = data.text.trim();
                    }
                    
                    console.log('Văn bản nhận dạng được (raw text):', recognizedText);
                    showNotification('Đã nhận dạng văn bản từ ảnh');
                    resolve(recognizedText); // Trả về văn bản
                    
                }).catch(err => {
                    console.error('Lỗi OCR Tesseract:', err);
                    showNotification('Lỗi khi nhận dạng văn bản từ ảnh.', 'error');
                    reject(err);
                });
            }
        };
        
        img.onerror = function() {
            console.error('Lỗi khi tải ảnh gốc cho processImageForOCR');
            showNotification('Không thể tải ảnh gốc để xử lý.', 'error');
            reject(new Error('Could not load original image'));
        };
        
        // Bắt đầu tải ảnh gốc
        console.log('Loading original image into Image object...');
        img.src = imageDataUrl;
    });
}

// Hàm kiểm tra trình duyệt
function checkBrowser() {
    const userAgent = navigator.userAgent.toLowerCase();
    const isIOS = /iphone|ipad|ipod/.test(userAgent);
    const isChrome = /chrome/.test(userAgent);
    const isAndroid = /android/.test(userAgent);
    
    // Chỉ cảnh báo khi là Chrome trên iOS
    if (isIOS && isChrome && !isAndroid) {
        showNotification('Vui lòng sử dụng Safari để có trải nghiệm tốt nhất. Chrome trên iOS có thể không hỗ trợ đầy đủ tính năng chụp ảnh.', 'warning');
        return false;
    }
    return true;
}

// Hàm xử lý ảnh được chụp trực tiếp
async function handleCaptureSelect(event) {
    console.log('Bắt đầu xử lý ảnh chụp...');
    
    // Kiểm tra trình duyệt trước khi xử lý
    if (!checkBrowser()) {
        document.getElementById('captureName').textContent = 'Chưa chụp';
        return;
    }
    
    const file = event.target.files[0];

    if (!file) {
        console.log('Không có file được chọn');
        document.getElementById('captureName').textContent = 'Chưa chụp';
        return;
    }

    document.getElementById('captureName').textContent = 'Đã chụp ảnh';
    showNotification('Đã chụp ảnh. Vui lòng chọn vùng cần xử lý...');

    const reader = new FileReader();

    reader.onload = function(e) {
        console.log('Đã đọc xong file ảnh');
        const imageDataUrl = e.target.result;
        showImageCropModal(imageDataUrl);
    };

    reader.onerror = function() {
        console.error('Lỗi khi đọc ảnh chụp');
        showNotification('Không thể đọc ảnh chụp. Vui lòng thử lại.', 'error');
    };

    reader.readAsDataURL(file);
}

// Hàm xử lý việc chọn file ảnh để OCR
async function handleImageSelect(event) {
    // Kiểm tra trình duyệt trước khi xử lý
    if (!checkBrowser()) {
        document.getElementById('imageName').textContent = 'Chưa có ảnh';
        return;
    }
    
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

// Hàm hiển thị modal chọn vùng ảnh
async function showImageCropModal(imageDataUrl) {
    console.log('Bắt đầu hiển thị modal...');
    const modal = document.getElementById('imageCropModal');
    const cropImage = document.getElementById('cropImage');
    
    if (!modal || !cropImage) {
        console.error('Không tìm thấy modal hoặc cropImage element');
        return;
    }
    
    // Hiển thị thông báo xử lý
    showNotification('Đang tiền xử lý ảnh...');

    try {
        // Xử lý ảnh và lấy Data URL của ảnh đã xử lý để hiển thị
        const processedImageDataUrl = await processImageForOCR(imageDataUrl, true);
        
        // Đặt src cho ảnh trong modal là ảnh đã xử lý
        cropImage.src = processedImageDataUrl;
        currentImage = cropImage;
        
        // Hiển thị modal
        modal.style.display = 'block';
        console.log('Modal đã được hiển thị với ảnh đã xử lý');
        
        // Xóa vùng chọn cũ nếu có
        if (selectionOverlay) {
            selectionOverlay.remove();
            selectionOverlay = null;
        }
        
        // Thêm sự kiện cho việc chọn vùng
        // Đảm bảo ảnh đã tải xong trước khi thêm listeners
        cropImage.onload = function() {
            console.log('Processed image loaded in modal');
            const container = document.querySelector('.image-container');
            
            if (!container) {
                console.error('Không tìm thấy image-container');
                return;
            }
            
            // Xóa các event listener cũ
            container.removeEventListener('mousedown', handleMouseDown);
            container.removeEventListener('mousemove', handleMouseMove);
            container.removeEventListener('mouseup', handleMouseUp);
            container.removeEventListener('touchstart', handleTouchStart);
            container.removeEventListener('touchmove', handleTouchMove);
            container.removeEventListener('touchend', handleTouchEnd);
            
            // Thêm sự kiện mới
            container.addEventListener('mousedown', handleMouseDown);
            container.addEventListener('mousemove', handleMouseMove);
            container.addEventListener('mouseup', handleMouseUp);
            container.addEventListener('touchstart', handleTouchStart, { passive: false });
            container.addEventListener('touchmove', handleTouchMove, { passive: false });
            container.addEventListener('touchend', handleTouchEnd);
            
            console.log('Đã thêm các event listener cho container');
            // Ẩn thông báo xử lý sau khi ảnh hiện lên
            // showNotification('Sẵn sàng chọn vùng ảnh'); // Có thể thêm thông báo này
        };

         cropImage.onerror = function() {
            console.error('Lỗi khi tải ảnh đã xử lý vào modal');
            showNotification('Không thể tải ảnh đã xử lý. Vui lòng thử lại.', 'error');
             // Đóng modal nếu ảnh không tải được
            modal.style.display = 'none';
        };

         // Nếu ảnh đã được cache, onload có thể không chạy, kiểm tra trạng thái complete
        if (cropImage.complete) {
             console.log('Processed image already complete in modal');
             // Gọi trực tiếp logic thêm event listeners nếu ảnh đã tải xong ngay lập tức
             const container = document.querySelector('.image-container');
             if (container) {
                 container.removeEventListener('mousedown', handleMouseDown);
                 container.removeEventListener('mousemove', handleMouseMove);
                 container.removeEventListener('mouseup', handleMouseUp);
                 container.removeEventListener('touchstart', handleTouchStart);
                 container.removeEventListener('touchmove', handleTouchMove);
                 container.removeEventListener('touchend', handleTouchEnd);
                 
                 container.addEventListener('mousedown', handleMouseDown);
                 container.addEventListener('mousemove', handleMouseMove);
                 container.addEventListener('mouseup', handleMouseUp);
                 container.addEventListener('touchstart', handleTouchStart, { passive: false });
                 container.addEventListener('touchmove', handleTouchMove, { passive: false });
                 container.addEventListener('touchend', handleTouchEnd);
                 console.log('Đã thêm các event listener cho container (ảnh cached)');
             }
        }
        
    } catch (error) {
        console.error('Lỗi trong quá trình tiền xử lý ảnh để hiển thị modal:', error);
        showNotification('Lỗi khi chuẩn bị ảnh để chọn vùng.', 'error');
        // Đóng modal nếu có lỗi
        modal.style.display = 'none';
    }
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
    console.log('Touch start event');
    e.preventDefault(); // Ngăn chặn scroll khi touch
    const touch = e.touches[0];
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    // Kiểm tra nếu touch vào handle để resize
    if (e.target.classList.contains('selection-handle')) {
        console.log('Touch vào handle để resize');
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
        console.log('Touch vào vùng chọn để di chuyển');
        isMoving = true;
        startX = touch.clientX;
        startY = touch.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        return;
    }
    
    // Tạo vùng chọn mới
    console.log('Bắt đầu tạo vùng chọn mới');
    isSelecting = true;
    startX = touch.clientX - rect.left;
    startY = touch.clientY - rect.top;
    
    // Xóa vùng chọn cũ nếu có
    if (selectionOverlay) {
        selectionOverlay.remove();
        selectionOverlay = null;
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
    console.log('Đã tạo xong vùng chọn mới');
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
async function cropImage() {
    console.log('cropImage function called');
    if (!selectionOverlay || !currentImage) {
        console.log('No selection overlay or current image found');
        return;
    }
    
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    
    const rect = selectionOverlay.getBoundingClientRect();
    const imageRect = currentImage.getBoundingClientRect();
    
    // Tính toán tỷ lệ giữa ảnh gốc và ảnh hiển thị trong modal
    // currentImage.naturalWidth/Height là kích thước ảnh sau tiền xử lý hiển thị trong modal
    // imageRect.width/height là kích thước DOM element của ảnh trong modal
    const scaleX = currentImage.naturalWidth / imageRect.width;
    const scaleY = currentImage.naturalHeight / imageRect.height;
    
    // Tính toán vị trí và kích thước thực tế của vùng cắt trên ảnh đã tiền xử lý
    const cropX = (rect.left - imageRect.left) * scaleX;
    const cropY = (rect.top - imageRect.top) * scaleY;
    const cropWidth = rect.width * scaleX;
    const cropHeight = rect.height * scaleY;

    console.log(`Crop coordinates and size: x=${cropX}, y=${cropY}, width=${cropWidth}, height=${cropHeight}`);
    
    // Kiểm tra kích thước cắt hợp lệ
    if (cropWidth <= 0 || cropHeight <= 0) {
        console.log('Invalid crop size');
        showNotification('Vùng chọn không hợp lệ.', 'warning');
        return;
    }

    canvas.width = cropWidth;
    canvas.height = cropHeight;
    
    // Vẽ phần ảnh đã cắt từ ảnh đã tiền xử lý lên canvas mới
    ctx.drawImage(
        currentImage, // Sử dụng ảnh đã tiền xử lý đang hiển thị trong modal
        cropX, cropY, cropWidth, cropHeight,
        0, 0, cropWidth, cropHeight
    );
    
    // Chuyển đổi canvas thành Data URL
    const croppedImageDataUrl = canvas.toDataURL('image/jpeg');
    console.log('Cropped image Data URL obtained', croppedImageDataUrl.substring(0, 50) + '...');
    
    // Đóng modal
    document.getElementById('imageCropModal').style.display = 'none';
    console.log('Image crop modal closed');
    
    // Xử lý ảnh đã cắt (thực hiện OCR)
    try {
        console.log('Calling processImageForOCR with cropped image');
        const recognizedText = await processImageForOCR(croppedImageDataUrl);
        console.log('processImageForOCR resolved. Recognized text:', recognizedText);
        
        // Thay thế các ký tự xuống dòng bằng khoảng trắng và trim
        const cleanedText = recognizedText.replace(/[\r\n]+/g, ' ').trim();
        console.log('Cleaned text:', cleanedText);
        
        // Điền văn bản đã xử lý vào ô tìm kiếm
        const searchInput = document.getElementById('searchInput');
        if (searchInput) {
            searchInput.value = cleanedText;
            console.log('Search input value set:', searchInput.value);
            showNotification('Đã nhận dạng và điền văn bản vào ô tìm kiếm');
            // Tự động gọi tìm kiếm sau khi nhận dạng thành công
            searchExcelData(cleanedText);
            console.log('searchExcelData called');
        } else {
            console.error('Search input element not found!');
            showNotification('Lỗi: Không tìm thấy ô tìm kiếm.', 'error');
        }
        
    } catch (error) {
        console.error('Lỗi xử lý ảnh cắt:', error);
        showNotification('Lỗi khi xử lý ảnh cắt.', 'error');
    }
}

// Thêm kiểm tra trình duyệt khi trang được tải
window.onload = () => {
    // Kiểm tra trình duyệt
    checkBrowser();
    
    // Đăng ký sự kiện cho nút chọn file Excel
    document.getElementById('fileInput').addEventListener('change', handleFileSelect);
    
    // Đăng ký sự kiện cho nút chọn file ảnh
    document.getElementById('imageInput').addEventListener('change', handleImageSelect);
    
    // Thêm trình lắng nghe sự kiện paste
    document.body.addEventListener('paste', handlePaste);
    
    // Thêm xử lý paste cho ô tìm kiếm
    document.getElementById('searchInput').addEventListener('paste', handleSearchInputPaste);
    
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
        searchExcelData(searchTerm);
    };
    
    // Tìm kiếm khi nhấn phím Enter
    document.getElementById('searchInput').addEventListener('keypress', function(event) {
        if (event.key === 'Enter') {
            event.preventDefault();
            const searchTerm = document.getElementById('searchInput').value;
            searchExcelData(searchTerm);
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

// Hàm xử lý sự kiện paste
function handlePaste(event) {
    const items = (event.clipboardData || event.originalEvent.clipboardData).items;
    let blob = null;

    for (const item of items) {
        // Tìm kiếm item có kiểu là image
        if (item.type.indexOf('image') === 0) {
            blob = item.getAsFile();
            break;
        }
    }

    if (blob) {
        console.log('Đã dán ảnh từ clipboard');
        showNotification('Đã phát hiện ảnh từ clipboard. Đang xử lý...');
        
        // Đọc blob ảnh dưới dạng Data URL
        const reader = new FileReader();
        reader.onload = function(e) {
            const imageDataUrl = e.target.result;
            // Gọi hàm xử lý OCR chung
            processImageForOCR(imageDataUrl);
        };
        reader.onerror = function() {
            console.error('Lỗi khi đọc blob ảnh từ clipboard');
            showNotification('Không thể đọc ảnh từ clipboard.', 'error');
        };
        reader.readAsDataURL(blob);
    } else {
        console.log('Không có ảnh trong clipboard');
    }
}

// Hàm xử lý paste cho ô tìm kiếm
async function handleSearchInputPaste(e) {
    console.log('Paste event triggered on search input');
    const items = (e.clipboardData || window.clipboardData).items;
    let blob = null;

    for (const item of items) {
        // Tìm kiếm item có kiểu là image
        if (item.type.indexOf('image') === 0) {
            blob = item.getAsFile();
            console.log('Image found in clipboard');
            break;
        }
    }

    if (blob) {
        e.preventDefault(); // Ngăn chặn paste mặc định của ảnh
        console.log('Processing pasted image...');
        showNotification('Đã phát hiện ảnh. Đang nhận dạng văn bản...');
        
        // Đọc blob ảnh dưới dạng Data URL
        const reader = new FileReader();
        reader.onload = async function(e) {
            const imageDataUrl = e.target.result;
            try {
                // Gọi hàm xử lý OCR và nhận văn bản trả về
                const recognizedText = await processImageForOCR(imageDataUrl);
                // Thay thế các ký tự xuống dòng bằng khoảng trắng và trim
                const cleanedText = recognizedText.replace(/[\r\n]+/g, ' ').trim();
                // Điền văn bản đã xử lý vào ô tìm kiếm
                document.getElementById('searchInput').value = cleanedText;
                showNotification('Đã nhận dạng và điền văn bản vào ô tìm kiếm');
                // Tự động gọi tìm kiếm sau khi nhận dạng thành công
                searchExcelData(cleanedText);
                
            } catch (error) {
                console.error('Lỗi xử lý ảnh dán:', error);
                showNotification('Lỗi khi xử lý ảnh dán.', 'error');
            }
        };
        reader.onerror = function() {
            console.error('Lỗi khi đọc ảnh từ clipboard');
            showNotification('Không thể đọc ảnh từ clipboard.', 'error');
        };
        reader.readAsDataURL(blob);
    } else {
        console.log('No image found, processing as text paste');
        // Nếu không phải ảnh, xử lý paste văn bản bình thường
        // Lấy văn bản, thay thế xuống dòng bằng khoảng trắng và trim
        const pastedText = (e.clipboardData || window.clipboardData).getData('text');
        const cleanedText = pastedText.replace(/[\r\n]+/g, ' ').trim();

        // Cho phép hành vi paste mặc định để chèn văn bản đã làm sạch
        // hoặc chèn thủ công nếu cần kiểm soát chính xác vị trí con trỏ
        
        // Tạm thời bỏ preventDefault() ở đây để xem paste mặc định có hoạt động không
        // e.preventDefault(); 
        e.target.value = cleanedText; // Ghi đè giá trị input
        console.log('Pasted text:', cleanedText);

        // Tự động gọi tìm kiếm sau khi paste văn bản
        searchExcelData(cleanedText);
    }
}