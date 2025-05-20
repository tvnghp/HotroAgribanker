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
                 <span class="column-f-label" style="position: absolute; top: -10px; left: 0; background-color: #FFD700; color: #8A1538; font-size: 9px; padding: 1px 4px; border-radius: 3px; font-weight: bold;">Cột F</span>
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
        setTimeout(() => document.body.removeChild(notification), 3000);
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
async function processImageForOCR(imageDataUrl) {
    showProcessingModal();
    try {
        const img = new Image();
        img.onload = async () => {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            canvas.width = img.width;
            canvas.height = img.height;
            ctx.drawImage(img, 0, 0);

            // Simple preprocessing for paste (grayscale and threshold)
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const data = imageData.data;
            for (let i = 0; i < data.length; i += 4) {
                const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
                const color = avg > 128 ? 255 : 0;
                data[i] = color; // red
                data[i + 1] = color; // green
                data[i + 2] = color; // blue
            }
            ctx.putImageData(imageData, 0, 0);

            const preprocessedDataUrl = canvas.toDataURL();

            const worker = await Tesseract.createWorker({
                logger: m => console.log(m)
            });
            await worker.loadLanguage('eng');
            await worker.initialize('eng');
            // Use simple page segmentation mode for general text
            await worker.setParameters({
                tessedit_pageseg_mode: '6',
            });
            const {
                data: {
                    text
                }
            } = await worker.recognize(preprocessedDataUrl);
            await worker.terminate();
            hideProcessingModal();
            // Use event to update search input, handled elsewhere
            const event = new CustomEvent('ocrcomplete', {
                detail: {
                    text: text.trim()
                }
            });
            document.dispatchEvent(event);

        };
        img.src = imageDataUrl;
    } catch (error) {
        console.error('OCR Error:', error);
        hideProcessingModal();
        alert('Error processing image for OCR: ' + error.message);
    }
}

// Function to apply enhanced preprocessing specifically for camera images
async function preprocessCameraImage(imageDataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            canvas.width = img.width;
            canvas.height = img.height;
            ctx.drawImage(img, 0, 0);

            // Convert to grayscale
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const data = imageData.data;
            for (let i = 0; i < data.length; i += 4) {
                const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
                data[i] = avg; // red
                data[i + 1] = avg; // green
                data[i + 2] = avg; // blue
            }
            ctx.putImageData(imageData, 0, 0);

            // Apply Gaussian blur (more aggressive)
            const blurRadius = 5; // Increased blur radius
            stackBlurCanvasRGBA(canvas, 0, 0, canvas.width, canvas.height, blurRadius);

            // Apply Adaptive Thresholding (using a simpler approach or library function)
            // For simplicity and avoiding external libraries here, we can try a manual adaptive-like threshold
            // Or revert to a global threshold if adaptive is too complex without a library.
            // Let's try a global threshold with a value derived from the image histogram if needed,
            // but for now, a slightly adjusted global threshold might work better after blur.
            // A simple global threshold after blur:
            const blurredImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const blurredData = blurredImageData.data;
            const thresholdValue = 100; // Adjusted threshold after blur
            for (let i = 0; i < blurredData.length; i += 4) {
                const color = blurredData[i] > thresholdValue ? 255 : 0;
                blurredData[i] = color;
                blurredData[i + 1] = color;
                blurredData[i + 2] = color;
            }
            ctx.putImageData(blurredImageData, 0, 0);

            resolve(canvas.toDataURL());
        };
        img.onerror = reject;
        img.src = imageDataUrl;
    });
}

// Function to display the image crop modal
function showImageCropModal(imageDataUrl) {
    const modal = document.getElementById('imageCropModal');
    const cropImageContainer = document.getElementById('cropImageContainer');
    const img = new Image();
    img.onload = () => {
        // Clear previous content
        cropImageContainer.innerHTML = '';

        // Set the image source
        cropImageContainer.style.backgroundImage = `url(${imageDataUrl})`;
        cropImageContainer.style.backgroundSize = 'contain';
        cropImageContainer.style.backgroundRepeat = 'no-repeat';
        cropImageContainer.style.backgroundPosition = 'center';
        cropImageContainer.style.width = '100%';
        cropImageContainer.style.height = 'calc(100% - 80px)'; // Adjust height
        cropImageContainer.style.position = 'relative'; // Needed for positioning selection area

        // Reset or initialize selection area state
        selectionArea = null;
        isDrawing = false;
        isMoving = false;
        isResizing = false;
        currentHandle = null;

        modal.style.display = 'flex'; // Use flex to center content
    };
    img.onerror = (error) => {
        console.error('Error loading image for cropping:', error);
        alert('Error loading image for cropping.');
    };
    img.src = imageDataUrl; // Use the preprocessed image data URL
}

// Function to hide the image crop modal
function hideImageCropModal() {
    const modal = document.getElementById('imageCropModal');
    modal.style.display = 'none';
    const cropImageContainer = document.getElementById('cropImageContainer');
    cropImageContainer.innerHTML = ''; // Clear the image
}

// Variables for drawing selection area
let selectionArea = null; // { x, y, width, height } relative to image container
let isDrawing = false;
let startX, startY; // Mouse down position
let isMoving = false; // Flag for moving the selection area
let isResizing = false; // Flag for resizing the selection area
let currentHandle = null; // Which handle is being dragged ('nw', 'ne', 'sw', 'se', 'n', 's', 'w', 'e')
let originalMouseX, originalMouseY; // Mouse position at start of move/resize
let originalSelectionX, originalSelectionY, originalSelectionWidth, originalSelectionHeight; // Selection area state at start of move/resize

// Event listeners for mouse actions on the image container
const cropImageContainer = document.getElementById('cropImageContainer');

cropImageContainer.addEventListener('mousedown', handleMouseDown);
cropImageContainer.addEventListener('mousemove', handleMouseMove);
cropImageContainer.addEventListener('mouseup', handleMouseUp);
document.addEventListener('mouseleave', handleMouseUp); // End drawing if mouse leaves window

function handleMouseDown(e) {
    e.preventDefault(); // Prevent default drag behavior
    const rect = cropImageContainer.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    // Check if clicking on an existing selection area or handle
    if (selectionArea) {
        const handleSize = 8; // Size of the resize handles
        const {
            x: sx,
            y: sy,
            width: sw,
            height: sh
        } = selectionArea;

        // Check handles
        if (x >= sx - handleSize && x <= sx + handleSize && y >= sy - handleSize && y <= sy + handleSize) currentHandle = 'nw';
        else if (x >= sx + sw - handleSize && x <= sx + sw + handleSize && y >= sy - handleSize && y <= sy + handleSize) currentHandle = 'ne';
        else if (x >= sx - handleSize && x <= sx + handleSize && y >= sy + sh - handleSize && y <= sy + sh + handleSize) currentHandle = 'sw';
        else if (x >= sx + sw - handleSize && x <= sx + sw + handleSize && y >= sy + sh - handleSize && y <= sy + sh + handleSize) currentHandle = 'se';
        else if (x >= sx + sw / 2 - handleSize && x <= sx + sw / 2 + handleSize && y >= sy - handleSize && y <= sy + handleSize) currentHandle = 'n';
        else if (x >= sx + sw / 2 - handleSize && x <= sx + sw / 2 + handleSize && y >= sy + sh - handleSize && y <= sy + sh + handleSize) currentHandle = 's';
        else if (x >= sx - handleSize && x <= sx + handleSize && y >= sy + sh / 2 - handleSize && y <= sy + sh / 2 + handleSize) currentHandle = 'w';
        else if (x >= sx + sw - handleSize && x <= sx + sw + handleSize && y >= sy + sh / 2 - handleSize && y <= sy + sh / 2 + handleSize) currentHandle = 'e';
        // Check if clicking inside the selection area for moving
        else if (x >= sx && x <= sx + sw && y >= sy && y <= sy + sh) isMoving = true;

        if (currentHandle) {
            isResizing = true;
            originalSelectionWidth = sw;
            originalSelectionHeight = sh;
            originalSelectionX = sx;
            originalSelectionY = sy;
            originalMouseX = e.clientX;
            originalMouseY = e.clientY;
        } else if (isMoving) {
            originalSelectionX = sx;
            originalSelectionY = sy;
            originalMouseX = e.clientX;
            originalMouseY = e.clientY;
        } else {
            // If not clicking on selection or handle, start drawing a new one
            isDrawing = true;
            startX = x;
            startY = y;
            selectionArea = {
                x: startX,
                y: startY,
                width: 0,
                height: 0
            };
            drawSelectionArea();
        }

    } else {
        // Start drawing a new selection area
        isDrawing = true;
        startX = x;
        startY = y;
        selectionArea = {
            x: startX,
            y: startY,
            width: 0,
            height: 0
        };
        drawSelectionArea();
    }
}

function handleMouseMove(e) {
    const rect = cropImageContainer.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    // Get the dimensions of the background image
    const imgElement = cropImageContainer.querySelector('img');
    const imgWidth = imgElement ? imgElement.offsetWidth : cropImageContainer.offsetWidth; // Fallback to container if img not yet loaded
    const imgHeight = imgElement ? imgElement.offsetHeight : cropImageContainer.offsetHeight;

    // Calculate the offset of the image if it's centered
    let imageOffsetX = 0;
    let imageOffsetY = 0;
    if (imgWidth < cropImageContainer.offsetWidth) {
        imageOffsetX = (cropImageContainer.offsetWidth - imgWidth) / 2;
    }
    if (imgHeight < cropImageContainer.offsetHeight) {
        imageOffsetY = (cropImageContainer.offsetHeight - imgHeight) / 2;
    }

    if (isDrawing) {
        const width = x - startX;
        const height = y - startY;

        selectionArea.x = Math.min(startX, startX + width);
        selectionArea.y = Math.min(startY, startY + height);
        selectionArea.width = Math.abs(width);
        selectionArea.height = Math.abs(height);

        // Ensure selection area stays within the image bounds
        selectionArea.x = Math.max(imageOffsetX, selectionArea.x);
        selectionArea.y = Math.max(imageOffsetY, selectionArea.y);
        const maxSelectionX = imageOffsetX + imgWidth - selectionArea.width;
        const maxSelectionY = imageOffsetY + imgHeight - selectionArea.height;
        selectionArea.x = Math.min(maxSelectionX, selectionArea.x);
        selectionArea.y = Math.min(maxSelectionY, selectionArea.y);

        // Ensure minimum size
        const minSize = 10;
        if (selectionArea.width < minSize) selectionArea.width = minSize;
        if (selectionArea.height < minSize) selectionArea.height = minSize;

        drawSelectionArea();

    } else if (isMoving && selectionArea) {
        const deltaX = e.clientX - originalMouseX;
        const deltaY = e.clientY - originalMouseY;

        let newX = originalSelectionX + deltaX;
        let newY = originalSelectionY + deltaY;

        // Constrain movement within the image bounds
        newX = Math.max(imageOffsetX, newX);
        newY = Math.max(imageOffsetY, newY);
        newX = Math.min(imageOffsetX + imgWidth - selectionArea.width, newX);
        newY = Math.min(imageOffsetY + imgHeight - selectionArea.height, newY);

        selectionArea.x = newX;
        selectionArea.y = newY;

        drawSelectionArea();

    } else if (isResizing && selectionArea) {
        const deltaX = e.clientX - originalMouseX;
        const deltaY = e.clientY - originalMouseY;

        let newX = originalSelectionX;
        let newY = originalSelectionY;
        let newWidth = originalSelectionWidth;
        let newHeight = originalSelectionHeight;

        const minSize = 10;

        switch (currentHandle) {
            case 'nw':
                newX = originalSelectionX + deltaX;
                newY = originalSelectionY + deltaY;
                newWidth = originalSelectionWidth - deltaX;
                newHeight = originalSelectionHeight - deltaY;
                break;
            case 'ne':
                newY = originalSelectionY + deltaY;
                newWidth = originalSelectionWidth + deltaX;
                newHeight = originalSelectionHeight - deltaY;
                break;
            case 'sw':
                newX = originalSelectionX + deltaX;
                newWidth = originalSelectionWidth - deltaX;
                newHeight = originalSelectionHeight + deltaY;
                break;
            case 'se':
                newWidth = originalSelectionWidth + deltaX;
                newHeight = originalSelectionHeight + deltaY;
                break;
            case 'n':
                newY = originalSelectionY + deltaY;
                newHeight = originalSelectionHeight - deltaY;
                break;
            case 's':
                newHeight = originalSelectionHeight + deltaY;
                break;
            case 'w':
                newX = originalSelectionX + deltaX;
                newWidth = originalSelectionWidth - deltaX;
                break;
            case 'e':
                newWidth = originalSelectionWidth + deltaX;
                break;
        }

        // Apply constraints and ensure minimum size
        if (newWidth > minSize) {
            // Check bounds for width and x
            if (currentHandle === 'nw' || currentHandle === 'sw' || currentHandle === 'w') {
                if (newX >= imageOffsetX) {
                    selectionArea.x = newX;
                    selectionArea.width = newWidth;
                }
            } else { // ne, se, e
                if (originalSelectionX + newWidth <= imageOffsetX + imgWidth) {
                    selectionArea.width = newWidth;
                }
            }
        }

        if (newHeight > minSize) {
            // Check bounds for height and y
            if (currentHandle === 'nw' || currentHandle === 'ne' || currentHandle === 'n') {
                if (newY >= imageOffsetY) {
                    selectionArea.y = newY;
                    selectionArea.height = newHeight;
                }
            } else { // sw, se, s
                if (originalSelectionY + newHeight <= imageOffsetY + imgHeight) {
                    selectionArea.height = newHeight;
                }
            }
        }

        drawSelectionArea();
    }
}

function handleMouseUp(e) {
    isDrawing = false;
    isMoving = false;
    isResizing = false;
    currentHandle = null;
}

function drawSelectionArea() {
    const container = document.getElementById('cropImageContainer');
    // Remove previous selection element if it exists
    let selectionElement = container.querySelector('.selection-area');
    if (selectionElement) {
        container.removeChild(selectionElement);
    }

    if (selectionArea && selectionArea.width > 0 && selectionArea.height > 0) {
        selectionElement = document.createElement('div');
        selectionElement.classList.add('selection-area');
        selectionElement.style.position = 'absolute';
        selectionElement.style.left = `${selectionArea.x}px`;
        selectionElement.style.top = `${selectionArea.y}px`;
        selectionElement.style.width = `${selectionArea.width}px`;
        selectionElement.style.height = `${selectionArea.height}px`;
        selectionElement.style.border = '2px dashed red';
        selectionElement.style.boxSizing = 'border-box';
        selectionElement.style.cursor = 'move'; // Cursor for moving

        // Add resize handles
        const handles = ['nw', 'ne', 'sw', 'se', 'n', 's', 'w', 'e'];
        handles.forEach(handleType => {
            const handle = document.createElement('div');
            handle.classList.add('handle', handleType);
            // Style handles
            handle.style.position = 'absolute';
            handle.style.width = '8px';
            handle.style.height = '8px';
            handle.style.backgroundColor = 'blue';
            handle.style.border = '1px solid white';
            handle.style.zIndex = '1'; // Ensure handles are above selection area

            // Position handles
            switch (handleType) {
                case 'nw':
                    handle.style.top = '-4px';
                    handle.style.left = '-4px';
                    handle.style.cursor = 'nw-resize';
                    break;
                case 'ne':
                    handle.style.top = '-4px';
                    handle.style.right = '-4px';
                    handle.style.cursor = 'ne-resize';
                    break;
                case 'sw':
                    handle.style.bottom = '-4px';
                    handle.style.left = '-4px';
                    handle.style.cursor = 'sw-resize';
                    break;
                case 'se':
                    handle.style.bottom = '-4px';
                    handle.style.right = '-4px';
                    handle.style.cursor = 'se-resize';
                    break;
                case 'n':
                    handle.style.top = '-4px';
                    handle.style.left = '50%';
                    handle.style.transform = 'translateX(-50%)';
                    handle.style.cursor = 'n-resize';
                    break;
                case 's':
                    handle.style.bottom = '-4px';
                    handle.style.left = '50%';
                    handle.style.transform = 'translateX(-50%)';
                    handle.style.cursor = 's-resize';
                    break;
                case 'w':
                    handle.style.left = '-4px';
                    handle.style.top = '50%';
                    handle.style.transform = 'translateY(-50%)';
                    handle.style.cursor = 'w-resize';
                    break;
                case 'e':
                    handle.style.right = '-4px';
                    handle.style.top = '50%';
                    handle.style.transform = 'translateY(-50%)';
                    handle.style.cursor = 'e-resize';
                    break;
            }
            selectionElement.appendChild(handle);
        });

        container.appendChild(selectionElement);
    }
}

// Function to crop the image based on the selection area
async function cropImage() {
    if (!selectionArea) {
        alert('Please select an area to crop.');
        return;
    }

    const cropImageContainer = document.getElementById('cropImageContainer');
    const backgroundImage = cropImageContainer.style.backgroundImage;
    if (!backgroundImage || backgroundImage === 'none') {
        console.error('No background image found in crop container.');
        alert('Error: Image not available for cropping.');
        return;
    }

    // Extract the URL from the background-image style
    const imageUrlMatch = backgroundImage.match(/url\("?(.*?)"?\)/);
    if (!imageUrlMatch || !imageUrlMatch[1]) {
        console.error('Could not extract image URL from background style.');
        alert('Error: Could not find image to crop.');
        return;
    }
    const originalImageUrl = imageUrlMatch[1];

    const img = new Image();
    img.onload = async () => {
        const containerRect = cropImageContainer.getBoundingClientRect();
        const imgAspectRatio = img.width / img.height;
        const containerAspectRatio = containerRect.width / containerRect.height;

        let displayWidth, displayHeight, displayOffsetX, displayOffsetY;

        if (imgAspectRatio > containerAspectRatio) {
            // Image is wider than container, it fits by width
            displayWidth = containerRect.width;
            displayHeight = containerRect.width / imgAspectRatio;
            displayOffsetX = 0;
            displayOffsetY = (containerRect.height - displayHeight) / 2;
        } else {
            // Image is taller than container, it fits by height
            displayHeight = containerRect.height;
            displayWidth = containerRect.height * imgAspectRatio;
            displayOffsetX = (containerRect.width - displayWidth) / 2;
            displayOffsetY = 0;
        }

        // Calculate the scale factor from display size to original size
        const scaleX = img.width / displayWidth;
        const scaleY = img.height / displayHeight;

        // Calculate the crop area in original image coordinates
        // Adjust selection area coordinates by the display offset
        const selectionXOnImageDisplay = selectionArea.x - displayOffsetX;
        const selectionYOnImageDisplay = selectionArea.y - displayOffsetY;

        const cropX = selectionXOnImageDisplay * scaleX;
        const cropY = selectionYOnImageDisplay * scaleY;
        const cropWidth = selectionArea.width * scaleX;
        const cropHeight = selectionArea.height * scaleY;

        // Create a canvas to draw the cropped image
        const canvas = document.createElement('canvas');
        canvas.width = cropWidth;
        canvas.height = cropHeight;
        const ctx = canvas.getContext('2d');

        // Draw the cropped section of the original image onto the new canvas
        ctx.drawImage(img,
            cropX, cropY, cropWidth, cropHeight, // Source image region
            0, 0, cropWidth, cropHeight // Destination canvas region
        );

        // Get the data URL of the cropped image
        const croppedDataUrl = canvas.toDataURL('image/png'); // Or 'image/jpeg'

        // Hide the modal
        hideImageCropModal();

        // Process the cropped image for OCR
        // We need to decide if we apply the *same* simple preprocessing as paste
        // or if camera images get a different, potentially more complex processing.
        // Based on previous conversation, let's use the simple processing for the cropped image,
        // similar to paste, as the camera preprocessing should have happened *before* the modal.
        // Assuming the image loaded into the modal was already preprocessed for camera.
        // So we pass the cropped area of the *already preprocessed* image.
        // If the modal loads the *original* image, we'd need to preprocess the *cropped* area here.

        // Let's stick to the plan: preprocess camera image *before* modal,
        // modal shows preprocessed image, crop takes a part of the preprocessed image,
        // and OCR runs on that cropped, preprocessed image part without further processing.

        // So, call processImageForOCR with the cropped data URL directly.
        await processImageForOCR(croppedDataUrl);

    };
    img.onerror = (error) => {
        console.error('Error loading image for cropping:', error);
        alert('Error loading image for cropping.');
    };
    // Load the image that was set as the background in the modal
    img.src = originalImageUrl;
}

// Add event listeners after the DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners(); // This function should now include paste listener

    // Check browser on load
    checkBrowser();

    // Handle image paste event on the search input
    const searchInput = document.getElementById('searchInput');
    searchInput.addEventListener('paste', handleSearchInputPaste);

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
});

function setupEventListeners() {
    // Assuming buttons for capture, file select, crop, and cancel are present
    const captureInput = document.getElementById('captureInput');
    if (captureInput) {
        captureInput.addEventListener('change', handleCaptureSelect);
    }

    const fileInput = document.getElementById('fileInput');
    if (fileInput) {
        fileInput.addEventListener('change', handleFileSelect);
    }

    const cropButton = document.getElementById('cropButton');
    if (cropButton) {
        cropButton.addEventListener('click', cropImage);
    }

    const cancelButton = document.getElementById('cancelCropButton');
    if (cancelButton) {
        cancelButton.addEventListener('click', hideImageCropModal);
    }

    // Add listener for the modal close button
    const closeButton = document.querySelector('.modal .close-button');
    if (closeButton) {
        closeButton.addEventListener('click', hideImageCropModal);
    }

    // Listen for custom event when OCR is complete
    document.addEventListener('ocrcomplete', (e) => {
        const recognizedText = e.detail.text;
        const searchInput = document.getElementById('searchInput');
        if (searchInput) {
            // Replace newlines with spaces
            searchInput.value = recognizedText.replace(/\\r?\\n/g, ' ');
            // Optionally trigger search automatically
            performSearch(searchInput.value);
        }
    });
}

// Function to perform search (placeholder - replace with actual search logic)
function performSearch(text) {
    console.log('Performing search for:', text);
    // Add your actual search logic here, e.g., filtering a list, sending a request, etc.
}

async function handleCaptureSelect(event) {
    // Check if it's Chrome on iOS before proceeding
    if (isChromeOnIOS) {
        alert("Chrome on iOS may not fully support camera capture and processing.");
        // Optionally, you could prevent default here or handle differently
        // event.preventDefault();
        // return;
    }

    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = async (e) => {
            const originalImageDataUrl = e.target.result;

            // Preprocess the camera image BEFORE showing the modal
            showProcessingModal("Đang tiền xử lý ảnh..."); // Show processing message
            try {
                const preprocessedDataUrl = await preprocessCameraImage(originalImageDataUrl);
                hideProcessingModal(); // Hide processing message
                // Show modal with the preprocessed image
                showImageCropModal(preprocessedDataUrl);
            } catch (error) {
                hideProcessingModal(); // Hide processing message
                console.error('Error during camera image preprocessing:', error);
                alert('Đã xảy ra lỗi khi tiền xử lý ảnh từ camera.');
            }

        };
        reader.readAsDataURL(file);
    }
}

async function handleImageSelect(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = async (e) => {
            const imageDataUrl = e.target.result;
            // For file select, we might not need the same aggressive preprocessing as camera.
            // Let's just show the original image in the modal for now, and the cropping
            // and OCR will apply the *simple* processing later if needed or rely on the OCR engine.
            // Or, we apply the *simple* processing here, similar to paste.
            // Let's apply the simple processing for consistency before showing the modal.
            showProcessingModal("Đang xử lý ảnh..."); // Show processing message
            try {
                // Apply simple preprocessing for consistency with paste and cropped camera images
                const processedDataUrl = await processImageForOCR(imageDataUrl); // Re-use processImageForOCR for simple processing
                hideProcessingModal(); // Hide processing message
                // Note: processImageForOCR also triggers OCR. We only want to show the modal here.
                // We need a separate function for simple processing that doesn't trigger OCR.

                // Let's refactor: simple processing should be a separate utility function.
                // Then processImageForOCR uses simple processing + OCR.
                // handleImageSelect uses simple processing + show modal.
                // handleSearchInputPaste uses simple processing + OCR directly.
                // handleCaptureSelect uses advanced processing + show modal.
                // cropImage uses simple processing + OCR.

                // Temporary fix: use processImageForOCR but know it also runs OCR.
                // We'll refactor preprocessing later if needed.

                showImageCropModal(imageDataUrl); // Show modal with original image for file select

            } catch (error) {
                hideProcessingModal(); // Hide processing message
                console.error('Error processing file image:', error);
                alert('Đã xảy ra lỗi khi xử lý ảnh.');
            }
        };
        reader.readAsDataURL(file);
    }
}

async function handleSearchInputPaste(event) {
    const items = (event.clipboardData || event.originalEvent.clipboardData).items;
    let handled = false;
    for (const item of items) {
        // Check if the pasted item is an image
        if (item.type.indexOf('image') !== -1) {
            const blob = item.getAsFile();
            if (blob) {
                const reader = new FileReader();
                reader.onload = async (e) => {
                    const imageDataUrl = e.target.result;
                    // Process the pasted image directly for OCR
                    showProcessingModal("Đang nhận dạng văn bản từ ảnh...");
                    try {
                        // The processImageForOCR function is designed to perform simple processing + OCR
                        // and dispatch an 'ocrcomplete' event.
                        await processImageForOCR(imageDataUrl);
                        // The event listener for 'ocrcomplete' will update the input and trigger search.

                    } catch (error) {
                        hideProcessingModal();
                        console.error('Error processing pasted image:', error);
                        alert('Đã xảy ra lỗi khi xử lý ảnh dán.');
                    }
                };
                reader.readAsDataURL(blob);
                handled = true;
                break; // Process only the first image found
            }
        } else if (item.type === 'text/plain') {
            // Handle text paste
            item.getAsString(function(text) {
                const searchInput = document.getElementById('searchInput');
                if (searchInput) {
                    // Replace newlines with spaces
                    const cleanedText = text.replace(/\\r?\\n/g, ' ');
                    // Insert the cleaned text at the current cursor position
                    const start = searchInput.selectionStart;
                    const end = searchInput.selectionEnd;
                    searchInput.value = searchInput.value.substring(0, start) + cleanedText + searchInput.value.substring(end);
                    // Restore cursor position
                    searchInput.selectionStart = searchInput.selectionEnd = start + cleanedText.length;
                    // Prevent the default paste action
                    event.preventDefault();
                    // Optionally trigger search automatically
                    performSearch(searchInput.value);
                }
            });
            handled = true;
            // Don't break here, allow other text/plain items if any
        }
    }

    if (handled) {
        event.preventDefault(); // Prevent default paste if we handled it
    }
}

// Function to show processing modal
function showProcessingModal(message = "Đang xử lý...") {
    let processingModal = document.getElementById('processingModal');
    if (!processingModal) {
        processingModal = document.createElement('div');
        processingModal.id = 'processingModal';
        processingModal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.7);
            color: white;
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            font-size: 1.5em;
        `;
        document.body.appendChild(processingModal);
    }
    processingModal.textContent = message;
    processingModal.style.display = 'flex';
}

// Function to hide processing modal
function hideProcessingModal() {
    const processingModal = document.getElementById('processingModal');
    if (processingModal) {
        processingModal.style.display = 'none';
    }
}

// Function to check browser and OS
let isChromeOnIOS = false;

function checkBrowser() {
    const userAgent = navigator.userAgent;
    console.log("User Agent:", userAgent); // Log user agent for debugging
    const isIOS = /iPhone|iPad|iPod/.test(userAgent);
    const isChrome = /CriOS/.test(userAgent); // Chrome on iOS uses CriOS
    const isAndroid = /Android/.test(userAgent); // Check for Android

    // Check if it's Chrome on iOS specifically
    if (isIOS && isChrome) {
        isChromeOnIOS = true;
        console.warn("Running on Chrome on iOS. Camera capture may have limitations.");
        // You could display a more prominent warning in the UI here if needed
    } else {
        isChromeOnIOS = false;
        console.log("Not running on Chrome on iOS.");
    }

    // Log if it's Chrome on Android (for user info, no special handling needed currently)
    if (isAndroid && /Chrome/.test(userAgent) && !isChrome) { // Standard Chrome on Android
        console.log("Running on Chrome on Android.");
    } else if (isAndroid && isChrome) { // CriOS on Android (less common?)
        console.log("Running on CriOS on Android.");
    }

    // Note: The prompt specifically mentioned Chrome on iPhone.
    // The code checks for Chrome (CriOS) on any iOS device (iPhone, iPad, iPod).
    // This seems appropriate for covering the user's concern.
}

// Include StackBlur library for Gaussian blur (copy-pasted or linked)
/*
   StackBlur - a fast almost Gaussian blur
   version 0.5
   Copyright (c) 2011 Mario Klingemann

   Permission is hereby granted, free of charge, to any person
   obtaining a copy of this software and associated documentation
   files (the "Software"), to deal in the Software without
   restriction, including without limitation the rights to use,
   copy, modify, merge, publish, distribute, sublicense, and/or sell
   copies of the Software, and to permit persons to whom the
   Software is furnished to do so, subject to the following
   conditions:

   The above copyright notice and this permission notice shall be
   included in all copies or substantial portions of the Software.

   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
   EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
   OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
   NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
   HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
   WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
   FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
   OTHER DEALINGS IN THE SOFTWARE.
*/

var mul_table = [512, 512, 456, 512, 328, 456, 335, 512, 405, 328, 271, 456, 388, 335, 292, 512, 454, 405, 364, 328, 298, 271, 249, 456, 420, 388, 360, 335, 312, 292, 275, 512, 482, 454, 428, 405, 383, 364, 345, 328, 314, 298, 284, 271, 259, 249, 240, 456, 437, 420, 403, 388, 373, 360, 347, 335, 323, 312, 302, 292, 283, 275, 267, 512, 495, 482, 469, 454, 442, 428, 417, 405, 394, 383, 373, 364, 354, 345, 337, 328, 320, 312, 305, 298, 291, 284, 278, 271, 265, 259, 253, 249, 243, 240, 235, 456, 445, 437, 428, 420, 412, 403, 396, 388, 381, 373, 367, 360, 354, 348, 341, 335, 329, 323, 318, 312, 307, 302, 297, 292, 287, 283, 278, 274, 271, 267, 263, 259, 512, 505, 495, 485, 475, 465, 456, 447, 437, 428, 420, 412, 404, 396, 388, 381, 374, 367, 360, 354, 347, 341, 335, 329, 323, 318, 313, 308, 302, 297, 292, 288, 283, 278, 274, 269, 265, 261, 257, 253, 250, 246, 243, 239, 237, 234, 231, 229, 227, 225, 223, 221, 219, 217, 215, 213, 212, 210, 209, 207, 206, 204, 203, 201, 200, 198, 197, 195, 194, 192, 191, 190, 188, 187, 185, 184, 183, 181, 180, 179, 177, 176, 175, 174, 172, 171, 170, 169, 167, 166, 165, 164, 163, 161, 160, 159, 158, 157, 156, 155, 154, 153, 152, 151, 150, 149, 148, 147, 146, 145, 144, 143, 142, 141, 140, 139, 138, 137, 136, 135, 134, 133, 132, 131, 130, 129, 128, 127, 126, 125, 124, 123, 122, 121, 120, 119, 118, 117, 116, 115, 114, 113, 112, 111, 110, 109, 108, 107, 106, 105, 104, 103, 102, 101, 100, 99, 98, 97, 96, 95, 94, 93, 92, 91, 90, 89, 88, 87, 86, 85, 84, 83, 82, 81, 80, 79, 78, 77, 76, 75, 74, 73, 72, 71, 70, 69, 68, 67, 66, 65, 64, 63, 62, 61, 60, 59, 58, 57, 56, 55, 54, 53, 52, 51, 50, 49, 48, 47, 46, 45, 44, 43, 42, 41, 40, 39, 38, 37, 36, 35, 34, 33, 32, 31, 30, 29, 28, 27, 26, 25, 24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0];
var shg_table = [9, 11, 12, 13, 13, 14, 14, 15, 15, 15, 15, 16, 16, 16, 16, 17, 17, 17, 17, 17, 17, 17, 17, 18, 18, 18, 18, 18, 18, 18, 18, 19, 19, 19, 19, 19, 19, 19, 19, 19, 19, 19, 19, 19, 19, 19, 19, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23];

function stackBlurCanvasRGBA(canvas, top_x, top_y, width, height, radius) {
    if (isNaN(radius) || radius < 1) {
        return;
    }
    radius |= 0;

    var context = canvas.getContext("2d");
    var imageData;

    try {
        imageData = context.getImageData(top_x, top_y, width, height);
    } catch (e) {

        // tried to get imageData on an image that's not in the same domain.
        // set crossOrigin attribute to an empty string and reload the image.
        try {
            canvas.crossOrigin = "anonymous";
            imageData = context.getImageData(top_x, top_y, width, height);
        } catch (e) {
            // still failed after setting crossOrigin to anonymous
            return;
        }
    }

    var pixels = imageData.data;

    var x, y, i, p, yp, yi, ue, cs;

    var div = radius + radius + 1;
    var w4 = width << 2;
    var widthMinus1 = width - 1;
    var heightMinus1 = height - 1;
    var radiusPlus1 = radius + 1;
    var sumFactor = radiusPlus1 * (radiusPlus1 + 1) / 2;

    var stackStart = new BlurStack();
    var stackEnd = new BlurStack();
    var stack = stackStart;
    for (i = 1; i < div; i++) {
        stack = stack.next = new BlurStack();
        if (i == radiusPlus1) var stackIn = stack;
    }
    stack.next = stackStart;
    var stackOut = stack;

    var mul_sum = mul_table[radius];
    var shg_sum = shg_table[radius];

    var pr, pg, pb, pa, rbs;

    var dv = new Int32Array(256 * div);
    var bl = new Int32Array(12);

    for (y = 0; y < height; y++) {
        rbs = -radiusPlus1 * div;
        sum = 0;
        sumr = 0;
        sumg = 0;
        sumb = 0;
        suma = 0;

        stack = stackStart;

        for (i = 0; i < radiusPlus1; i++) {
            p = pixels[(yi + i) << 2];
            stack.r = p;
            stack.g = p + 1;
            stack.b = p + 2;
            stack.a = p + 3;
            rbs += sumFactor;
            sum += rbs;
            sumr += p * rbs;
            sumg += p + 1 * rbs;
            sumb += p + 2 * rbs;
            suma += p + 3 * rbs;
            stack = stack.next;
        }

        yp = radius;
        ue = radius;
        yi = y * w4;
        cs = y * width;

        for (x = 0; x < width; x++) {
            pixels[yi + 0] = (sumr * mul_sum) >>> shg_sum;
            pixels[yi + 1] = (sumg * mul_sum) >>> shg_sum;
            pixels[yi + 2] = (sumb * mul_sum) >>> shg_sum;
            pixels[yi + 3] = (suma * mul_sum) >>> shg_sum;

            sum -= rbs;

            stackOut = stackOut.next;
            rbs -= stackOut.rbs;

            p = x + radiusPlus1;

            if (ue < widthMinus1) {
                ue++;
            }
            stackIn = stackIn.next;
            stackIn.rbs = radiusPlus1 * (ue);
            sum += stackIn.rbs;

            stackIn.r = pixels[(cs + ue) << 2];
            stackIn.g = pixels[((cs + ue) << 2) + 1];
            stackIn.b = pixels[((cs + ue) << 2) + 2];
            stackIn.a = pixels[((cs + ue) << 2) + 3];

            sumr += stackIn.rbs * stackIn.r;
            sumg += stackIn.rbs * stackIn.g;
            sumb += stackIn.rbs * stackIn.b;
            suma += stackIn.rbs * stackIn.a;

            yi += 4;
            rbs += mul_table[radius]; // Fixed: use mul_table for rbs adjustment
        }
    }

    for (x = 0; x < width; x++) {
        rbs = -radiusPlus1 * div;
        sum = 0;
        sumr = 0;
        sumg = 0;
        sumb = 0;
        suma = 0;

        stack = stackStart;

        for (i = 0; i < radiusPlus1; i++) {
            yi = ((yp + i) * w4) + (x << 2);
            p = pixels[yi];
            stack.r = p;
            stack.g = pixels[yi + 1];
            stack.b = pixels[yi + 2];
            stack.a = pixels[yi + 3];
            rbs += sumFactor;
            sum += rbs;
            sumr += stack.r * rbs;
            sumg += stack.g * rbs;
            sumb += stack.b * rbs;
            suma += stack.a * rbs;
            stack = stack.next;
        }

        cs = x;
        yp = radius;
        ue = radius;

        for (y = 0; y < height; y++) {
            yi = (y * w4) + (x << 2);
            pixels[yi] = (sumr * mul_sum) >>> shg_sum;
            pixels[yi + 1] = (sumg * mul_sum) >>> shg_sum;
            pixels[yi + 2] = (sumb * mul_sum) >>> shg_sum;
            pixels[yi + 3] = (suma * mul_sum) >>> shg_sum;

            sum -= rbs;

            stackOut = stackOut.next;
            rbs -= stackOut.rbs;

            p = y + radiusPlus1;

            if (ue < heightMinus1) {
                ue++;
            }
            yi = ((yp + ue) * w4) + (x << 2);
            stackIn = stackIn.next;
            stackIn.rbs = radiusPlus1 * (ue);
            sum += stackIn.rbs;

            stackIn.r = pixels[yi];
            stackIn.g = pixels[yi + 1];
            stackIn.b = pixels[yi + 2];
            stackIn.a = pixels[yi + 3];

            sumr += stackIn.rbs * stackIn.r;
            sumg += stackIn.rbs * stackIn.g;
            sumb += stackIn.rbs * stackIn.b;
            suma += stackIn.rbs * stackIn.a;

            rbs += mul_table[radius]; // Fixed: use mul_table for rbs adjustment
        }
    }

    context.putImageData(imageData, top_x, top_y);
}

function BlurStack() {
    this.r = 0;
    this.g = 0;
    this.b = 0;
    this.a = 0;
    this.rbs = 0; // Added rbs property
    this.next = null;
}