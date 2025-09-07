// Canvas Configuration - Fixed dimensions for PPTX compatibility
const CANVAS_CONFIGS = {
    msb: {
        name: 'MSB',
        width: 342,
        height: 350,
        color: '#3b82f6',
        dimensions: '9.05cm × 9.25cm'
    },
    mccb: {
        name: 'MCCB',
        width: 320,
        height: 346,
        color: '#10b981',
        dimensions: '4.22cm × 4.57cm'
    },
    tpsld: {
        name: 'TP_SLD',
        width: 320,
        height: 346,
        color: '#8b5cf6',
        dimensions: '4.22cm × 4.57cm'
    },
    tpmccbcompartment: {
        name: 'TP_MCCB_COMPARTMENT',
        width: 320,
        height: 346,
        color: '#6366f1',
        dimensions: '4.22cm × 4.57cm'
    },
    tptappingloc: {
        name: 'TP_TAPPING_LOC',
        width: 320,
        height: 346,
        color: '#f43f5e',
        dimensions: '4.22cm × 4.57cm'
    },
    tprouting1: {
        name: 'TP_ROUTING_1',
        widthCm: 8.85,
        heightCm: 10.45,
        color: '#f59e0b',
        responsive: true
    },
    tprouting2: {
        name: 'TP_ROUTING_2',
        widthCm: 8.85,
        heightCm: 10.45,
        color: '#14b8a6',
        responsive: true
    },
    tprouting3: {
        name: 'TP_ROUTING_3',
        widthCm: 17.74,
        heightCm: 9.28,
        color: '#84cc16',
        responsive: true
    }
};

// Logo Configuration
const LOGOS = [
    { name: 'Schneider', src: 'https://via.placeholder.com/60x40/0066cc/ffffff?text=SE' },
    { name: 'ABB', src: 'https://via.placeholder.com/60x40/ff0000/ffffff?text=ABB' },
    { name: 'Siemens', src: 'https://via.placeholder.com/60x40/00a651/ffffff?text=SIE' },
    // Add more logos as needed
];

// Global State
let canvases = {};
let canvasStates = {};
let currentCanvasType = null;
let currentTool = 'draw';
let uploadedImages = {};

// Initialize Application
document.addEventListener('DOMContentLoaded', function() {
    initializeUploadSections();
    initializeCanvasSelector();
    initializeLogoSection();
    initializeToolbar();
    calculateResponsiveDimensions();
    initializeCanvases();
    
    // Set default canvas
    switchCanvas('msb');
    
    // Handle window resize for responsive canvases
    window.addEventListener('resize', debounce(handleResize, 300));
});

// Initialize Upload Sections
function initializeUploadSections() {
    const uploadSection = document.getElementById('uploadSection');
    
    Object.entries(CANVAS_CONFIGS).forEach(([key, config]) => {
        const uploadCard = createUploadCard(key, config);
        uploadSection.appendChild(uploadCard);
    });
}

// Create Upload Card
function createUploadCard(canvasType, config) {
    const card = document.createElement('div');
    card.className = 'upload-card';
    
    card.innerHTML = `
        <h3 class="text-lg font-semibold text-gray-800 mb-4 text-center">${config.name} Upload</h3>
        <div class="upload-zone" onclick="document.getElementById('${canvasType}Input').click()">
            <svg class="upload-icon" style="color: ${config.color}" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
            </svg>
            <p class="mb-2 text-sm text-gray-500">
                <span class="font-semibold">Click to upload ${config.name}</span>
            </p>
            <p class="text-xs text-gray-500">PNG, JPG or JPEG</p>
            ${config.dimensions ? `<p class="text-xs mt-1" style="color: ${config.color}">Dimensions: ${config.dimensions}</p>` : ''}
            <input id="${canvasType}Input" type="file" class="hidden" accept="image/*" onchange="handleImageUpload('${canvasType}', this)" />
        </div>
        <div id="${canvasType}Preview" class="preview-section">
            <p class="text-sm text-gray-600 mb-2">Preview:</p>
            <img id="${canvasType}PreviewImg" src="" alt="${config.name} Preview" class="preview-img">
            <button onclick="removeImage('${canvasType}')" class="mt-2 px-3 py-1 bg-red-500 text-white rounded text-sm hover:bg-red-600">Remove</button>
        </div>
    `;
    
    return card;
}

// Initialize Canvas Selector
function initializeCanvasSelector() {
    const selector = document.getElementById('canvasSelector');
    selector.className = 'canvas-selector-grid';
    
    Object.entries(CANVAS_CONFIGS).forEach(([key, config]) => {
        const btn = document.createElement('button');
        btn.className = 'canvas-selector-btn';
        btn.textContent = config.name;
        btn.onclick = () => switchCanvas(key);
        btn.id = `${key}Selector`;
        selector.appendChild(btn);
    });
}

// Initialize Logo Section
function initializeLogoSection() {
    const logoSection = document.getElementById('logoSection');
    logoSection.className = 'logo-grid';
    
    LOGOS.forEach(logo => {
        const btn = document.createElement('div');
        btn.className = 'logo-btn';
        btn.onclick = () => addLogo(logo.src);
        btn.innerHTML = `<img src="${logo.src}" alt="${logo.name}" title="${logo.name}">`;
        logoSection.appendChild(btn);
    });
}

// Initialize Toolbar
function initializeToolbar() {
    document.getElementById('drawBtn').onclick = () => setTool('draw');
    document.getElementById('lineBtn').onclick = () => setTool('line');
    document.getElementById('eraseBtn').onclick = () => setTool('erase');
    document.getElementById('textBtn').onclick = () => setTool('text');
    document.getElementById('clearBtn').onclick = clearCurrentCanvas;
    document.getElementById('resetBtn').onclick = resetAll;
    document.getElementById('exportBtn').onclick = exportToPPTX;
}

// Calculate Responsive Dimensions for Routing Canvases
function calculateResponsiveDimensions() {
    const maxWidth = Math.min(window.innerWidth - 64, 800);
    const maxHeight = Math.min(window.innerHeight * 0.6, 600);
    
    Object.entries(CANVAS_CONFIGS).forEach(([key, config]) => {
        if (config.responsive) {
            const ratio = config.widthCm / config.heightCm;
            let width, height;
            
            if (ratio > 1) {
                width = Math.min(maxWidth, maxHeight * ratio);
                height = width / ratio;
            } else {
                height = Math.min(maxHeight, maxWidth / ratio);
                width = height * ratio;
            }
            
            // Ensure minimum size
            if (width < 280) {
                width = 280;
                height = width / ratio;
            }
            
            config.width = Math.round(width);
            config.height = Math.round(height);
        }
    });
}

// Initialize Canvases
function initializeCanvases() {
    const canvasWrapper = document.getElementById('canvasWrapper');
    
    Object.entries(CANVAS_CONFIGS).forEach(([key, config]) => {
        // Create canvas container
        const container = document.createElement('div');
        container.id = `${key}Container`;
        container.className = `canvas-container canvas-${key}`;
        container.style.width = `${config.width}px`;
        container.style.height = `${config.height}px`;
        
        // Create canvas element
        const canvas = document.createElement('canvas');
        canvas.id = `${key}Canvas`;
        canvas.width = config.width;
        canvas.height = config.height;
        
        // Create label
        const label = document.createElement('div');
        label.className = `canvas-label label-${key}`;
        label.textContent = `${config.name} Canvas`;
        
        // Create legend box
        const legend = document.createElement('div');
        legend.id = `${key}Legend`;
        legend.className = 'legend-box';
        
        container.appendChild(canvas);
        container.appendChild(label);
        container.appendChild(legend);
        canvasWrapper.appendChild(container);
        
        // Initialize Fabric.js canvas
        const fabricCanvas = new fabric.Canvas(canvas, {
            width: config.width,
            height: config.height,
            backgroundColor: '#f8f9fa',
            enableRetinaScaling: false
        });
        
        canvases[key] = fabricCanvas;
        canvasStates[key] = {
            image: null,
            cropRect: null,
            originalImageData: null
        };
        
        // Add crop rectangle
        const cropRect = new fabric.Rect({
            left: 0,
            top: 0,
            width: config.width,
            height: config.height,
            fill: 'transparent',
            stroke: config.color,
            strokeWidth: 2,
            selectable: false,
            evented: false,
            excludeFromExport: true
        });
        
        fabricCanvas.add(cropRect);
        canvasStates[key].cropRect = cropRect;
        
        // Setup event listeners
        setupCanvasEvents(fabricCanvas);
    });
}

// Setup Canvas Events
function setupCanvasEvents(canvas) {
    canvas.on('mouse:down', handleMouseDown);
    canvas.on('mouse:move', handleMouseMove);
    canvas.on('mouse:up', handleMouseUp);
    canvas.on('path:created', handlePathCreated);
}

// Switch Canvas
function switchCanvas(canvasType) {
    // Hide all canvases
    Object.keys(CANVAS_CONFIGS).forEach(key => {
        document.getElementById(`${key}Container`).classList.remove('active');
        document.getElementById(`${key}Selector`).classList.remove('active');
    });
    
    // Show selected canvas
    document.getElementById(`${canvasType}Container`).classList.add('active');
    document.getElementById(`${canvasType}Selector`).classList.add('active');
    
    currentCanvasType = canvasType;
}

// Handle Image Upload
function handleImageUpload(canvasType, input) {
    const file = input.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        uploadedImages[canvasType] = e.target.result;
        loadImageToCanvas(canvasType, e.target.result);
        
        // Show preview
        const preview = document.getElementById(`${canvasType}Preview`);
        const previewImg = document.getElementById(`${canvasType}PreviewImg`);
        previewImg.src = e.target.result;
        preview.classList.add('active');
    };
    reader.readAsDataURL(file);
}

// Load Image to Canvas
function loadImageToCanvas(canvasType, imageSrc) {
    const canvas = canvases[canvasType];
    const state = canvasStates[canvasType];
    const config = CANVAS_CONFIGS[canvasType];
    
    if (state.image) {
        canvas.remove(state.image);
    }
    
    fabric.Image.fromURL(imageSrc, function(img) {
        // Scale image to fit canvas
        const scaleX = config.width / img.width;
        const scaleY = config.height / img.height;
        const scale = Math.min(scaleX, scaleY);
        
        img.set({
            left: config.width / 2,
            top: config.height / 2,
            originX: 'center',
            originY: 'center',
            scaleX: scale,
            scaleY: scale,
            selectable: true,
            moveCursor: 'move',
            hoverCursor: 'move'
        });
        
        canvas.add(img);
        canvas.sendToBack(img);
        canvas.bringToFront(state.cropRect);
        
        state.image = img;
        state.originalImageData = imageSrc;
        canvas.renderAll();
    });
}

// Remove Image
function removeImage(canvasType) {
    const canvas = canvases[canvasType];
    const state = canvasStates[canvasType];
    
    if (state.image) {
        canvas.remove(state.image);
        state.image = null;
        state.originalImageData = null;
    }
    
    document.getElementById(`${canvasType}Preview`).classList.remove('active');
    document.getElementById(`${canvasType}Input`).value = '';
    delete uploadedImages[canvasType];
    canvas.renderAll();
}

// Tool Functions
function setTool(tool) {
    currentTool = tool;
    updateToolButtons();
    
    if (!currentCanvasType) return;
    
    const canvas = canvases[currentCanvasType];
    
    // Reset canvas modes
    canvas.isDrawingMode = false;
    canvas.selection = true;
    
    switch (tool) {
        case 'draw':
            canvas.isDrawingMode = true;
            canvas.freeDrawingBrush.width = 2;
            canvas.freeDrawingBrush.color = '#000000';
            break;
        case 'line':
            // Line drawing will be handled in mouse events
            break;
        case 'erase':
            canvas.isDrawingMode = true;
            canvas.freeDrawingBrush.width = 10;
            canvas.freeDrawingBrush.color = '#f8f9fa'; // Background color
            break;
        case 'text':
            // Text tool handled separately
            break;
    }
}

function updateToolButtons() {
    ['drawBtn', 'lineBtn', 'eraseBtn', 'textBtn'].forEach(btnId => {
        const btn = document.getElementById(btnId);
        btn.classList.remove('btn-active');
    });
    
    const activeBtn = document.getElementById(currentTool + 'Btn');
    if (activeBtn) {
        activeBtn.classList.add('btn-active');
    }
}

// Mouse Event Handlers
function handleMouseDown(options) {
    if (currentTool === 'text') {
        const pointer = this.getPointer(options.e);
        addText(pointer.x, pointer.y);
    }
}

function handleMouseMove(options) {
    // Handle mouse move for specific tools if needed
}

function handleMouseUp(options) {
    // Handle mouse up for specific tools if needed
}

function handlePathCreated(options) {
    // Handle path creation for drawing tools
}

// Add Text
function addText(x, y) {
    if (!currentCanvasType) return;
    
    const text = prompt('Enter text:');
    if (!text) return;
    
    const textObj = new fabric.Text(text, {
        left: x,
        top: y,
        fontFamily: 'Arial',
        fontSize: 16,
        fill: '#000000'
    });
    
    canvases[currentCanvasType].add(textObj);
}

// Add Logo
function addLogo(logoSrc) {
    if (!currentCanvasType) {
        alert('Please select a canvas first before adding logos.');
        return;
    }
    
    const canvas = canvases[currentCanvasType];
    const config = CANVAS_CONFIGS[currentCanvasType];
    
    fabric.Image.fromURL(logoSrc, function(img) {
        const maxSize = 80;
        const scale = Math.min(maxSize / img.width, maxSize / img.height);
        
        img.set({
            left: config.width / 2,
            top: config.height / 2,
            originX: 'center',
            originY: 'center',
            scaleX: scale,
            scaleY: scale,
            selectable: true
        });
        
        canvas.add(img);
        canvas.renderAll();
    });
}

// Clear Current Canvas
function clearCurrentCanvas() {
    if (!currentCanvasType) return;
    
    const canvas = canvases[currentCanvasType];
    const state = canvasStates[currentCanvasType];
    
    // Clear all objects except crop rectangle and image
    const objects = canvas.getObjects();
    objects.forEach(obj => {
        if (obj !== state.cropRect && obj !== state.image) {
            canvas.remove(obj);
        }
    });
    
    canvas.renderAll();
}

// Reset All
function resetAll() {
    if (confirm('Are you sure you want to reset everything?')) {
        Object.keys(CANVAS_CONFIGS).forEach(canvasType => {
            removeImage(canvasType);
            const canvas = canvases[canvasType];
            const state = canvasStates[canvasType];
            
            canvas.clear();
            canvas.add(state.cropRect);
            canvas.renderAll();
        });
    }
}

// Export to PPTX
function exportToPPTX() {
    const formData = new FormData();
    
    // Add canvas images to form data
    Object.entries(CANVAS_CONFIGS).forEach(([key, config]) => {
        const canvas = canvases[key];
        const dataURL = canvas.toDataURL('image/png');
        
        // Convert data URL to blob
        const arr = dataURL.split(',');
        const mime = arr[0].match(/:(.*?);/)[1];
        const bstr = atob(arr[1]);
        let n = bstr.length;
        const u8arr = new Uint8Array(n);
        
        while (n--) {
            u8arr[n] = bstr.charCodeAt(n);
        }
        
        const blob = new Blob([u8arr], { type: mime });
        formData.append(key, blob, `${key}.png`);
    });
    
    // Send to server
    fetch('/export', {
        method: 'POST',
        body: formData
    })
    .then(response => response.blob())
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'exported_presentation.pptx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    })
    .catch(error => {
        console.error('Export failed:', error);
        alert('Export failed. Please try again.');
    });
}

// Handle Window Resize
function handleResize() {
    calculateResponsiveDimensions();
    
    // Update routing canvases
    ['tprouting1', 'tprouting2', 'tprouting3'].forEach(canvasType => {
        const config = CANVAS_CONFIGS[canvasType];
        const container = document.getElementById(`${canvasType}Container`);
        const canvas = canvases[canvasType];
        
        if (container && canvas) {
            container.style.width = `${config.width}px`;
            container.style.height = `${config.height}px`;
            canvas.setDimensions({ width: config.width, height: config.height });
            canvas.renderAll();
        }
    });
}

// Utility Functions
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}