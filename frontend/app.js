// Constants
const API_BASE = 'http://localhost:5000/api';

// State Management
let currentTab = 'builder'; // 'builder' or 'template'
let blocks = [];
let activeImageBlockId = null;
let uploadedTemplateFile = null;
let uploadedTextTemplateContent = '';
let templateVariables = [];
let uploadedBaseTemplateFile = null;


// DOM Elements
const views = {
    builder: document.getElementById('view-builder'),
    template: document.getElementById('view-template')
};
const tabButtons = {
    builder: document.getElementById('btn-tab-builder'),
    template: document.getElementById('btn-tab-template')
};
const blocksContainer = document.getElementById('blocks-container');
const emptyState = document.getElementById('empty-state');
const uploadZone = document.getElementById('upload-zone');
const uploadedFileInfo = document.getElementById('uploaded-file-info');
const fileNameLabel = document.getElementById('file-name-label');
const templateFormCard = document.getElementById('template-form-card');
const templateDynamicForm = document.getElementById('template-dynamic-form');
const loadingOverlay = document.getElementById('loading-overlay');
const loadingText = document.getElementById('loading-text');
const meetingMetadata = document.getElementById('meeting-metadata');
const docStyleSelect = document.getElementById('doc-style');

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
    // Setup Drag and Drop for template upload
    setupDragAndDrop(uploadZone, handleTemplateFile);
    
    // Setup Page-wide Keyboard Paste Event
    setupGlobalPaste();
    
    // Update layout styling triggers
    handleStyleChange();
    
    showToast('WordDoc Workspace initialized!', 'info');
});

// Toast Notifications Helper
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let iconClass = 'fa-info-circle';
    if (type === 'success') iconClass = 'fa-check-circle';
    if (type === 'error') iconClass = 'fa-exclamation-triangle';
    
    toast.innerHTML = `<i class="fa-solid ${iconClass}"></i> <span>${message}</span>`;
    container.appendChild(toast);
    
    // Remove after 3.5 seconds
    setTimeout(() => {
        toast.style.animation = 'fadeOut 0.3s ease forwards';
        setTimeout(() => toast.remove(), 300);
    }, 3500);
}

// Switch Tabs
function switchTab(tab) {
    currentTab = tab;
    
    // Toggle active classes
    Object.keys(views).forEach(key => {
        if (key === tab) {
            views[key].classList.add('active');
            tabButtons[key].classList.add('active');
        } else {
            views[key].classList.remove('active');
            tabButtons[key].classList.remove('active');
        }
    });
    
    // Toggle sidebar config settings
    document.getElementById('builder-config').style.display = tab === 'builder' ? 'block' : 'none';
    document.getElementById('template-config').style.display = tab === 'template' ? 'block' : 'none';
}

// Show/Hide Meeting Notes fields
function handleStyleChange() {
    const style = docStyleSelect.value;
    if (style === 'meeting-notes') {
        meetingMetadata.style.display = 'block';
    } else {
        meetingMetadata.style.display = 'none';
    }
}

// Generate unique ID for blocks
function generateId() {
    return 'block-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
}

// Switch view empty state
function updateEmptyState() {
    if (blocks.length === 0) {
        emptyState.style.display = 'block';
    } else {
        emptyState.style.display = 'none';
    }
}

// Clear builder workspace
function clearWorkspace() {
    if (blocks.length === 0) return;
    if (confirm('Are you sure you want to clear your workspace? All pasted content will be lost.')) {
        blocks = [];
        activeImageBlockId = null;
        renderBlocks();
        showToast('Workspace cleared', 'info');
    }
}

// Create & Add a new block
function addBlock(type, initialData = null) {
    const newBlock = {
        id: generateId(),
        type: type,
        level: 1, // for headings
        content: '', // for heading, paragraph, code
        items: [''], // for bullet lists
        list_style: 'bullet',
        image_file: null, // for image block
        caption: '',
        width: 6.0
    };

    if (type === 'paragraph' && typeof initialData === 'string') {
        newBlock.content = initialData;
    } else if (type === 'image' && initialData instanceof File) {
        newBlock.image_file = initialData;
    } else if (type === 'heading' && typeof initialData === 'string') {
        newBlock.content = initialData;
    }

    blocks.push(newBlock);
    renderBlocks();
    
    // Smooth scroll to the bottom of the editor
    setTimeout(() => {
        const view = views.builder;
        view.scrollTo({ top: view.scrollHeight, behavior: 'smooth' });
        
        // Auto-focus the newly added block's input
        const blockEl = document.getElementById(newBlock.id);
        if (blockEl) {
            const input = blockEl.querySelector('input[type="text"], textarea');
            if (input) input.focus();
        }
    }, 50);

    return newBlock;
}

// Remove a specific block
function removeBlock(id) {
    blocks = blocks.filter(b => b.id !== id);
    if (activeImageBlockId === id) activeImageBlockId = null;
    renderBlocks();
}

// Reorder blocks (direction: -1 is Up, 1 is Down)
function moveBlock(id, direction) {
    const index = blocks.findIndex(b => b.id === id);
    if (index === -1) return;
    
    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= blocks.length) return;
    
    // Swap elements
    const temp = blocks[index];
    blocks[index] = blocks[targetIndex];
    blocks[targetIndex] = temp;
    
    renderBlocks();
    
    // Maintain focus highlight
    const blockEl = document.getElementById(id);
    if (blockEl) {
        blockEl.style.borderColor = 'var(--accent)';
        setTimeout(() => {
            blockEl.style.borderColor = 'var(--border-color)';
        }, 1000);
    }
}

// Global Paste Event Listener
function setupGlobalPaste() {
    window.addEventListener('paste', (e) => {
        // If we are in Template Mode, let the template paste handler deal with it
        if (currentTab === 'template') {
            return; 
        }
        
        // Don't intercept paste if user is typing in a textarea/input of an existing block,
        // EXCEPT if they are focusing an empty image slot.
        const active = document.activeElement;
        const isInput = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.isContentEditable);
        
        const items = (e.clipboardData || e.originalEvent.clipboardData).items;
        let imageFile = null;
        
        for (let i = 0; i < items.length; i++) {
            if (items[i].type.indexOf('image') !== -1) {
                const blob = items[i].getAsFile();
                imageFile = new File([blob], `screenshot_${Date.now()}.png`, { type: 'image/png' });
                break;
            }
        }
        
        if (imageFile) {
            e.preventDefault();
            
            // If the user has highlighted a specific empty image block, load image there!
            if (activeImageBlockId) {
                const block = blocks.find(b => b.id === activeImageBlockId);
                if (block && block.type === 'image') {
                    block.image_file = imageFile;
                    activeImageBlockId = null;
                    renderBlocks();
                    showToast('Screenshot loaded into selected slot!', 'success');
                    return;
                }
            }
            
            // Otherwise, create a new image block
            const newBlock = addBlock('image', imageFile);
            showToast('New screenshot block created!', 'success');
            
            // Auto focus caption
            setTimeout(() => {
                const blockEl = document.getElementById(newBlock.id);
                if (blockEl) {
                    const captionInput = blockEl.querySelector('.caption-input');
                    if (captionInput) captionInput.focus();
                }
            }, 100);
            
        } else if (!isInput) {
            // Paste plain text as a paragraph if user is not focusing any input field
            const text = e.clipboardData.getData('text');
            if (text && text.trim().length > 0) {
                e.preventDefault();
                addBlock('paragraph', text.trim());
                showToast('Text block added', 'info');
            }
        }
    });
}

// Drag & Drop Setup Utility
function setupDragAndDrop(element, callback) {
    ['dragenter', 'dragover'].forEach(eventName => {
        element.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            element.classList.add('drag-over');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        element.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            element.classList.remove('drag-over');
        }, false);
    });

    element.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files && files.length > 0) {
            callback(files[0]);
        }
    }, false);
}

// Render the Builder Blocks
function renderBlocks() {
    // Save current input values to state first
    saveAllInputsToState();
    
    // Clear container
    blocksContainer.innerHTML = '';
    updateEmptyState();
    
    blocks.forEach((block, index) => {
        const card = document.createElement('div');
        card.className = `block-card ${activeImageBlockId === block.id ? 'active-slot' : ''}`;
        card.id = block.id;
        
        // Reordering controls markup
        const upDisabled = index === 0 ? 'disabled style="opacity: 0.3; cursor: default;"' : '';
        const downDisabled = index === blocks.length - 1 ? 'disabled style="opacity: 0.3; cursor: default;"' : '';
        
        let blockContentHtml = '';
        
        // Render content based on block type
        switch (block.type) {
            case 'heading':
                blockContentHtml = `
                    <div class="heading-input-group">
                        <select onchange="updateBlockField('${block.id}', 'level', this.value)">
                            <option value="1" ${block.level == 1 ? 'selected' : ''}>H1 (Major)</option>
                            <option value="2" ${block.level == 2 ? 'selected' : ''}>H2 (Section)</option>
                            <option value="3" ${block.level == 3 ? 'selected' : ''}>H3 (Sub)</option>
                        </select>
                        <input type="text" value="${escapeHtml(block.content)}" placeholder="Heading text..." oninput="updateBlockField('${block.id}', 'content', this.value)">
                    </div>
                `;
                break;
                
            case 'paragraph':
                blockContentHtml = `
                    <textarea placeholder="Type or paste text paragraph..." oninput="updateBlockField('${block.id}', 'content', this.value)">${escapeHtml(block.content)}</textarea>
                `;
                break;
                
            case 'code':
                blockContentHtml = `
                    <textarea class="code-textarea" placeholder="Paste code snippet here..." oninput="updateBlockField('${block.id}', 'content', this.value)">${escapeHtml(block.content)}</textarea>
                `;
                break;
                
            case 'list':
                let listRows = '';
                block.items.forEach((item, i) => {
                    listRows += `
                        <div class="list-item-row">
                            <span class="list-bullet-dot"><i class="fa-solid fa-circle"></i></span>
                            <input type="text" value="${escapeHtml(item)}" placeholder="List item..." oninput="updateListField('${block.id}', ${i}, this.value)" onkeydown="handleListKeydown(event, '${block.id}', ${i})">
                            <button class="ctrl-btn btn-delete" onclick="removeListItem('${block.id}', ${i})" title="Remove item"><i class="fa-solid fa-xmark"></i></button>
                        </div>
                    `;
                });
                blockContentHtml = `
                    <div class="list-items-container" id="list-container-${block.id}">
                        ${listRows}
                    </div>
                    <button class="list-item-add" onclick="addListItem('${block.id}')">
                        <i class="fa-solid fa-plus"></i> Add Bullet Item
                    </button>
                `;
                break;
                
            case 'image':
                if (block.image_file) {
                    // Fetch object URL for preview
                    const imgUrl = URL.createObjectURL(block.image_file);
                    blockContentHtml = `
                        <div class="image-preview-container">
                            <img src="${imgUrl}" alt="Pasted image preview">
                            <button class="remove-img-btn" onclick="clearBlockImage('${block.id}')" title="Delete screenshot"><i class="fa-solid fa-trash-can"></i></button>
                        </div>
                        <div class="img-controls">
                            <div class="form-group">
                                <label>Description / Caption</label>
                                <input type="text" class="caption-input" value="${escapeHtml(block.caption)}" placeholder="e.g. Figure 1: App home screen" oninput="updateBlockField('${block.id}', 'caption', this.value)">
                            </div>
                            <div class="form-group" style="max-width: 120px;">
                                <label>Width (inches)</label>
                                <select onchange="updateBlockField('${block.id}', 'width', this.value)">
                                    <option value="4.0" ${block.width == 4.0 ? 'selected' : ''}>4.0"</option>
                                    <option value="5.0" ${block.width == 5.0 ? 'selected' : ''}>5.0"</option>
                                    <option value="6.0" ${block.width == 6.0 ? 'selected' : ''}>6.0" (Full)</option>
                                    <option value="7.0" ${block.width == 7.0 ? 'selected' : ''}>7.0"</option>
                                </select>
                            </div>
                        </div>
                    `;
                } else {
                    blockContentHtml = `
                        <div class="image-uploader-slot" onclick="focusImageSlot('${block.id}')" id="slot-${block.id}">
                            <input type="file" id="file-input-${block.id}" style="display: none;" accept="image/*" onchange="handleImageBlockUpload(event, '${block.id}')">
                            <i class="fa-solid fa-cloud-arrow-up"></i>
                            <p><strong>Click to browse</strong>, drag screenshot here, or click to focus and paste (Ctrl+V)</p>
                            <span>Supports PNG, JPG, WebP</span>
                        </div>
                    `;
                }
                break;
        }
        
        card.innerHTML = `
            <div class="block-controls">
                <span class="block-type-badge">${block.type}</span>
                <div class="reorder-btns">
                    <button class="ctrl-btn" onclick="moveBlock('${block.id}', -1)" ${upDisabled} title="Move Up"><i class="fa-solid fa-chevron-up"></i></button>
                    <button class="ctrl-btn" onclick="moveBlock('${block.id}', 1)" ${downDisabled} title="Move Down"><i class="fa-solid fa-chevron-down"></i></button>
                </div>
                <button class="ctrl-btn btn-delete" onclick="removeBlock('${block.id}')" title="Delete Block"><i class="fa-solid fa-trash-can"></i></button>
            </div>
            <div class="block-content">
                ${blockContentHtml}
            </div>
        `;
        
        blocksContainer.appendChild(card);
        
        // Setup individual drag & drop if image upload slot is showing
        if (block.type === 'image' && !block.image_file) {
            const slot = document.getElementById(`slot-${block.id}`);
            if (slot) {
                setupDragAndDrop(slot, (file) => {
                    block.image_file = file;
                    renderBlocks();
                });
            }
        }
    });
}

// Focus on an empty image slot, ready to receive Ctrl+V
function focusImageSlot(id) {
    activeImageBlockId = id;
    
    // Highlight visually
    document.querySelectorAll('.block-card').forEach(card => {
        card.classList.remove('active-slot');
    });
    const activeCard = document.getElementById(id);
    if (activeCard) activeCard.classList.add('active-slot');
    
    // Double click or simple click can also trigger browser browse
    const fileInput = document.getElementById(`file-input-${id}`);
    if (fileInput) {
        fileInput.click();
    }
}

// Handle traditional image file browsing inside blocks
function handleImageBlockUpload(e, id) {
    const file = e.target.files[0];
    if (file) {
        const block = blocks.find(b => b.id === id);
        if (block) {
            block.image_file = file;
            renderBlocks();
            showToast('Image uploaded', 'success');
        }
    }
}

// Clear loaded image and revert block back to empty slot
function clearBlockImage(id) {
    const block = blocks.find(b => b.id === id);
    if (block) {
        block.image_file = null;
        renderBlocks();
    }
}

// Save all block text inputs into state before re-rendering
function saveAllInputsToState() {
    blocks.forEach(block => {
        const card = document.getElementById(block.id);
        if (!card) return;
        
        if (block.type === 'heading') {
            const input = card.querySelector('input[type="text"]');
            const select = card.querySelector('select');
            if (input) block.content = input.value;
            if (select) block.level = parseInt(select.value);
        } else if (block.type === 'paragraph' || block.type === 'code') {
            const textarea = card.querySelector('textarea');
            if (textarea) block.content = textarea.value;
        } else if (block.type === 'image' && block.image_file) {
            const capInput = card.querySelector('.caption-input');
            const widthSelect = card.querySelector('select');
            if (capInput) block.caption = capInput.value;
            if (widthSelect) block.width = parseFloat(widthSelect.value);
        }
        // List items are saved directly via oninput, no need here
    });
}

// Updates block fields on key inputs
function updateBlockField(id, field, value) {
    const block = blocks.find(b => b.id === id);
    if (block) {
        if (field === 'level' || field === 'width') {
            block[field] = parseFloat(value);
        } else {
            block[field] = value;
        }
    }
}

// List specific updates
function updateListField(blockId, itemIndex, value) {
    const block = blocks.find(b => b.id === blockId);
    if (block && block.items) {
        block.items[itemIndex] = value;
    }
}

// List key bindings: Enter creates a new list row, Backspace on empty deletes row
function handleListKeydown(e, blockId, itemIndex) {
    const block = blocks.find(b => b.id === blockId);
    if (!block) return;
    
    if (e.key === 'Enter') {
        e.preventDefault();
        block.items.splice(itemIndex + 1, 0, '');
        renderBlocks();
        // Focus the new row
        setTimeout(() => {
            const card = document.getElementById(blockId);
            const inputs = card.querySelectorAll('.list-item-row input');
            if (inputs[itemIndex + 1]) inputs[itemIndex + 1].focus();
        }, 30);
    } else if (e.key === 'Backspace' && e.target.value === '' && block.items.length > 1) {
        e.preventDefault();
        block.items.splice(itemIndex, 1);
        renderBlocks();
        // Focus previous row
        setTimeout(() => {
            const card = document.getElementById(blockId);
            const inputs = card.querySelectorAll('.list-item-row input');
            const targetFocusIdx = itemIndex > 0 ? itemIndex - 1 : 0;
            if (inputs[targetFocusIdx]) inputs[targetFocusIdx].focus();
        }, 30);
    }
}

function addListItem(blockId) {
    const block = blocks.find(b => b.id === blockId);
    if (block && block.items) {
        block.items.push('');
        renderBlocks();
    }
}

function removeListItem(blockId, itemIndex) {
    const block = blocks.find(b => b.id === blockId);
    if (block && block.items && block.items.length > 1) {
        block.items.splice(itemIndex, 1);
        renderBlocks();
    }
}

// -------------------------------------------------------------
// TEMPLATE TAB FUNCTIONS
// -------------------------------------------------------------

function handleTemplateFile(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    const imageExtensions = ['png', 'jpg', 'jpeg', 'webp', 'bmp'];
    
    if (!['docx', 'md', 'txt'].concat(imageExtensions).includes(extension)) {
        showToast('Please select a valid template file (.docx, .md, .txt, or image)!', 'error');
        return;
    }
    
    uploadedTemplateFile = file;
    fileNameLabel.textContent = file.name;
    uploadedFileInfo.style.display = 'inline-flex';
    uploadZone.classList.add('file-loaded');
    
    if (extension === 'docx') {
        // Analyze .docx template via backend
        analyzeTemplateFile(file);
    } else if (imageExtensions.includes(extension)) {
        // Analyze image format template via OCR
        analyzeImageTemplateFile(file);
    } else {
        // Analyze .md / .txt template locally
        analyzeTextTemplateFile(file);
    }
}

function handleTemplateUpload(e) {
    const file = e.target.files[0];
    if (file) {
        handleTemplateFile(file);
    }
}

function clearTemplate(e) {
    e.stopPropagation();
    uploadedTemplateFile = null;
    templateVariables = [];
    uploadedFileInfo.style.display = 'none';
    uploadZone.classList.remove('file-loaded');
    templateFormCard.style.display = 'none';
    templateDynamicForm.innerHTML = '';
    document.getElementById('template-file').value = '';
    showToast('Template removed', 'info');
}

// Analyze template from the server
async function analyzeTemplateFile(file) {
    showLoading(true, 'Analyzing template file...');
    
    const formData = new FormData();
    formData.append('template', file);
    
    try {
        const response = await fetch(`${API_BASE}/analyze-template`, {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        if (response.ok) {
            templateVariables = data.variables;
            renderTemplateForm();
            showToast('Template parsed successfully!', 'success');
        } else {
            showToast(data.error || 'Failed to parse template', 'error');
            clearTemplate(new Event('clear'));
        }
    } catch (error) {
        showToast('Server connection failed. Is the backend running?', 'error');
        clearTemplate(new Event('clear'));
    } finally {
        showLoading(false);
    }
}

// Analyze Markdown/Text template locally
function analyzeTextTemplateFile(file) {
    showLoading(true, 'Analyzing template file...');
    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        uploadedTextTemplateContent = text;
        
        // Match standard placeholders: {{ variable }}
        const regex = /\{\{\s*([a-zA-Z0-9_-]+)\s*\}\}/g;
        const vars = new Set();
        let match;
        while ((match = regex.exec(text)) !== null) {
            vars.add(match[1]);
        }
        
        templateVariables = Array.from(vars).sort().map(varName => {
            const lowered = varName.toLowerCase();
            const isImg = ['image', 'img', 'pic', 'photo', 'screenshot', 'paste'].some(x => lowered.includes(x));
            return {
                name: varName,
                type: isImg ? 'image' : 'text',
                label: varName.replace(/_/g, ' ').replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase())
            };
        });
        
        renderTemplateForm();
        showLoading(false);
        showToast('Template layout loaded successfully!', 'success');
    };
    reader.onerror = function() {
        showLoading(false);
        showToast('Failed to read template file.', 'error');
        clearTemplate(new Event('clear'));
    };
    reader.readAsText(file);
}


// Render dynamic forms for templates
function renderTemplateForm() {
    templateDynamicForm.innerHTML = '';
    
    if (templateVariables.length === 0) {
        templateDynamicForm.innerHTML = `
            <div class="alert alert-info">
                <i class="fa-solid fa-circle-info"></i> No variables detected. Make sure to use standard Jinja tags in your word file (e.g. {{ title }}).
            </div>
        `;
        templateFormCard.style.display = 'block';
        return;
    }
    
    templateVariables.forEach(variable => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        if (variable.type === 'text') {
            formGroup.innerHTML = `
                <label for="tpl-field-${variable.name}">${variable.label}</label>
                <input type="text" id="tpl-field-${variable.name}" name="${variable.name}" placeholder="Enter ${variable.label.toLowerCase()}...">
            `;
        } else {
            // Image block in templates: allows pasting and instructions!
            formGroup.innerHTML = `
                <label>${variable.label} (Screenshot Image)</label>
                <div class="image-uploader-slot" id="tpl-slot-${variable.name}" onclick="focusTemplateImageSlot('${variable.name}')">
                    <input type="file" id="tpl-file-input-${variable.name}" name="${variable.name}" style="display: none;" accept="image/*" onchange="handleTemplateImageUpload(event, '${variable.name}')">
                    <div id="tpl-preview-container-${variable.name}" style="display: none; width: 100%;">
                        <div class="image-preview-container">
                            <img id="tpl-img-preview-${variable.name}" src="" alt="Preview">
                            <button type="button" class="remove-img-btn" onclick="clearTemplateImage(event, '${variable.name}')"><i class="fa-solid fa-trash-can"></i></button>
                        </div>
                    </div>
                    <div id="tpl-placeholder-${variable.name}">
                        <i class="fa-solid fa-image"></i>
                        <p><strong>Click to browse</strong>, or select slot and press <strong>Ctrl+V</strong> to paste screenshot</p>
                    </div>
                </div>
                <div class="tpl-image-instruction-box" style="margin-top: 8px;">
                    <input type="text" id="tpl-instruction-${variable.name}" placeholder="Type instructions or details to go with this screenshot..." style="width: 100%; padding: 8px 12px; background: rgba(0,0,0,0.25); border: 1px solid var(--border-color); border-radius: 8px; color: white; font-family: var(--font-sans); font-size: 13.5px; outline: none;">
                </div>
            `;
        }
        templateDynamicForm.appendChild(formGroup);
    });
    
    // Bind template-wide paste listeners on inputs
    bindTemplatePasteEvents();
    
    templateFormCard.style.display = 'block';
}

let activeTemplateImageField = null;

function focusTemplateImageSlot(name) {
    activeTemplateImageField = name;
    
    // Clear other slot outlines
    document.querySelectorAll('.image-uploader-slot').forEach(slot => {
        slot.style.borderColor = 'var(--border-color)';
    });
    
    const activeSlot = document.getElementById(`tpl-slot-${name}`);
    if (activeSlot) {
        activeSlot.style.borderColor = 'var(--accent)';
    }
    
    // Trigger browse
    document.getElementById(`tpl-file-input-${name}`).click();
}

function handleTemplateImageUpload(e, name) {
    const file = e.target.files[0];
    if (file) {
        loadTemplateImageFile(file, name);
    }
}

function loadTemplateImageFile(file, name) {
    const reader = new FileReader();
    reader.onload = (event) => {
        const preview = document.getElementById(`tpl-img-preview-${name}`);
        preview.src = event.target.result;
        
        document.getElementById(`tpl-preview-container-${name}`).style.display = 'block';
        document.getElementById(`tpl-placeholder-${name}`).style.display = 'none';
        
        // Cache the file object on the input element
        const input = document.getElementById(`tpl-file-input-${name}`);
        input.fileData = file;
        
        showToast('Image loaded successfully!', 'success');
    };
    reader.readAsDataURL(file);
}

function clearTemplateImage(e, name) {
    e.stopPropagation();
    document.getElementById(`tpl-preview-container-${name}`).style.display = 'none';
    document.getElementById(`tpl-placeholder-${name}`).style.display = 'block';
    
    const input = document.getElementById(`tpl-file-input-${name}`);
    input.value = '';
    input.fileData = null;
    
    if (activeTemplateImageField === name) {
        activeTemplateImageField = null;
    }
    
    // Reset border
    document.getElementById(`tpl-slot-${name}`).style.borderColor = 'var(--border-color)';
}

function bindTemplatePasteEvents() {
    window.addEventListener('paste', (e) => {
        if (currentTab !== 'template' || !activeTemplateImageField) return;
        
        const items = (e.clipboardData || e.originalEvent.clipboardData).items;
        let imageFile = null;
        
        for (let i = 0; i < items.length; i++) {
            if (items[i].type.indexOf('image') !== -1) {
                const blob = items[i].getAsFile();
                imageFile = new File([blob], `screenshot_${Date.now()}.png`, { type: 'image/png' });
                break;
            }
        }
        
        if (imageFile) {
            e.preventDefault();
            loadTemplateImageFile(imageFile, activeTemplateImageField);
            activeTemplateImageField = null;
        }
    });
}

// -------------------------------------------------------------
// LOADING & OVERLAY HELPERS
// -------------------------------------------------------------

function showLoading(show, text = 'Processing...') {
    loadingOverlay.style.display = show ? 'flex' : 'none';
    loadingText.textContent = text;
}

// -------------------------------------------------------------
// GENERATION CONTROL
// -------------------------------------------------------------

async function generateDocument() {
    if (currentTab === 'builder') {
        await generateFromBuilder();
    } else {
        await generateFromTemplate();
    }
}

// Create file from the workspace builder
async function generateFromBuilder() {
    saveAllInputsToState();
    
    if (blocks.length === 0) {
        showToast('Please add some content to your workspace before generating!', 'error');
        return;
    }
    
    // Check if there's at least a title
    const docTitle = document.getElementById('doc-title').value.trim() || 'My Document';
    
    showLoading(true, 'Building document structure...');
    
    const formData = new FormData();
    
    // Build JSON data
    const documentData = {
        title: docTitle,
        style: docStyleSelect.value,
        blocks: []
    };
    
    // For Meeting Notes style, package the metadata fields
    if (docStyleSelect.value === 'meeting-notes') {
        documentData.meeting_date = document.getElementById('meet-date').value || 'N/A';
        documentData.meeting_location = document.getElementById('meet-location').value || 'N/A';
        documentData.meeting_attendees = document.getElementById('meet-attendees').value || 'N/A';
        documentData.meeting_facilitator = document.getElementById('meet-facilitator').value || 'N/A';
    }
    
    // Map blocks to the output format
    blocks.forEach((block, index) => {
        const item = {
            type: block.type,
            level: block.level,
            content: block.content,
            caption: block.caption,
            width: block.width,
            list_style: block.list_style,
            items: block.items
        };
        
        if (block.type === 'image' && block.image_file) {
            const imageId = `image_${index}`;
            item.image_id = imageId;
            // Append file
            formData.append(imageId, block.image_file);
        }
        
        documentData.blocks.push(item);
    });
    
    formData.append('document_data', JSON.stringify(documentData));
    
    if (uploadedBaseTemplateFile) {
        formData.append('base_template', uploadedBaseTemplateFile);
    }
    
    try {

        const response = await fetch(`${API_BASE}/generate-from-scratch`, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            // Trigger browser download
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            
            // Format filename nicely
            const safeTitle = docTitle.replace(/[^a-z0-9]/gi, '_').toLowerCase();
            a.download = `${safeTitle || 'document'}_${docStyleSelect.value}.docx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            showToast('Document downloaded successfully!', 'success');
        } else {
            const data = await response.json();
            showToast(data.error || 'Failed to generate document', 'error');
        }
    } catch (error) {
        showToast('Error connecting to backend server', 'error');
    } finally {
        showLoading(false);
    }
}

// Generate file from an uploaded template
async function generateFromTemplate() {
    if (!uploadedTemplateFile) {
        showToast('Please upload a template file first!', 'error');
        return;
    }
    
    const extension = uploadedTemplateFile.name.split('.').pop().toLowerCase();
    if (extension !== 'docx') {
        // Compile Markdown/Text templates locally
        await generateFromTextTemplate();
        return;
    }

    
    showLoading(true, 'Assembling template fields...');
    
    const formData = new FormData();
    formData.append('template', uploadedTemplateFile);
    
    // Assemble text contexts and append image files
    const context = {};
    
    templateVariables.forEach(variable => {
        if (variable.type === 'text') {
            const input = document.getElementById(`tpl-field-${variable.name}`);
            context[variable.name] = input ? input.value : '';
        } else {
            const fileInput = document.getElementById(`tpl-file-input-${variable.name}`);
            const fileData = fileInput ? fileInput.fileData : null;
            if (fileData) {
                formData.append(variable.name, fileData);
                
                // Add screenshot description/instruction to context
                const instructionInput = document.getElementById(`tpl-instruction-${variable.name}`);
                if (instructionInput && instructionInput.value.trim() !== '') {
                    context[`${variable.name}_instruction`] = instructionInput.value.trim();
                }
            }
        }
    });
    
    formData.append('context', JSON.stringify(context));
    
    try {
        const response = await fetch(`${API_BASE}/generate-from-template`, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = `Filled_${uploadedTemplateFile.name}`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            showToast('Document downloaded successfully!', 'success');
        } else {
            const data = await response.json();
            showToast(data.error || 'Failed to generate template', 'error');
        }
    } catch (error) {
        showToast('Error connecting to backend server', 'error');
    } finally {
        showLoading(false);
    }
}

// HTML Escaping Utility
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    return text
        .toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// -------------------------------------------------------------
// OUTLINE FORMAT IMPORT FUNCTIONS (.MD / .TXT)
// -------------------------------------------------------------

function triggerOutlineUpload() {
    document.getElementById('outline-file-input').click();
}

async function handleOutlineUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const extension = file.name.split('.').pop().toLowerCase();
    const imageExtensions = ['png', 'jpg', 'jpeg', 'webp', 'bmp'];
    
    if (imageExtensions.includes(extension)) {
        try {
            const text = await performOCR(file);
            const importedBlocks = parseOutlineText(text);
            if (importedBlocks.length === 0) {
                showToast('No structured layout detected in screenshot.', 'warning');
                return;
            }
            
            if (blocks.length > 0 && !confirm('Importing this format outline will replace your current workspace blocks. Proceed?')) {
                e.target.value = '';
                return;
            }
            
            blocks = importedBlocks;
            activeImageBlockId = null;
            renderBlocks();
            showToast(`Extracted ${blocks.length} format blocks from screenshot!`, 'success');
        } catch (err) {
            showToast('OCR failed: ' + err.message, 'error');
        }
        e.target.value = '';
    } else {
        const reader = new FileReader();
        reader.onload = function(evt) {
            const text = evt.target.result;
            try {
                const importedBlocks = parseOutlineText(text);
                if (importedBlocks.length === 0) {
                    showToast('No structured format outline detected in file.', 'warning');
                    return;
                }
                
                if (blocks.length > 0 && !confirm('Importing this format outline will replace your current workspace blocks. Proceed?')) {
                    e.target.value = '';
                    return;
                }
                
                blocks = importedBlocks;
                activeImageBlockId = null;
                renderBlocks();
                showToast(`Successfully imported ${blocks.length} format blocks!`, 'success');
            } catch (err) {
                showToast('Failed to parse outline file: ' + err.message, 'error');
            }
            e.target.value = '';
        };
        reader.readAsText(file);
    }
}

async function performOCR(file) {
    showLoading(true, 'Initializing OCR engine...');
    try {
        const result = await Tesseract.recognize(
            file,
            'eng',
            { 
                logger: m => {
                    if (m.status === 'recognizing') {
                        const pct = Math.round(m.progress * 100);
                        loadingText.textContent = `Extracting format layout... ${pct}%`;
                    }
                } 
            }
        );
        return result.data.text;
    } catch (err) {
        console.error("Tesseract OCR Error: ", err);
        throw new Error(err.message || "OCR engine failed");
    } finally {
        showLoading(false);
    }
}

async function analyzeImageTemplateFile(file) {
    try {
        const text = await performOCR(file);
        uploadedTextTemplateContent = text;
        
        // Scan for placeholders: {{ variable }}
        const regex = /\{\{\s*([a-zA-Z0-9_-]+)\s*\}\}/g;
        const vars = new Set();
        let match;
        while ((match = regex.exec(text)) !== null) {
            vars.add(match[1]);
        }
        
        if (vars.size > 0) {
            templateVariables = Array.from(vars).sort().map(varName => {
                const lowered = varName.toLowerCase();
                const isImg = ['image', 'img', 'pic', 'photo', 'screenshot', 'paste'].some(x => lowered.includes(x));
                return {
                    name: varName,
                    type: isImg ? 'image' : 'text',
                    label: varName.replace(/_/g, ' ').replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase())
                };
            });
            renderTemplateForm();
            showToast('Detected template variables from screenshot OCR!', 'success');
        } else {
            // No placeholders, load recognized format blocks into builder
            const parsedBlocks = parseOutlineText(text);
            if (parsedBlocks.length > 0) {
                if (confirm('No {{ placeholders }} detected in screenshot. Load recognized format blocks into Document Workspace instead?')) {
                    blocks = parsedBlocks;
                    activeImageBlockId = null;
                    switchTab('builder');
                    renderBlocks();
                    showToast('Format structure loaded from screenshot successfully!', 'success');
                    clearTemplate(new Event('clear'));
                } else {
                    clearTemplate(new Event('clear'));
                }
            } else {
                showToast('Failed to extract structured text from screenshot.', 'warning');
                clearTemplate(new Event('clear'));
            }
        }
    } catch (err) {
        showToast('OCR recognition failed: ' + err.message, 'error');
        clearTemplate(new Event('clear'));
    }
}

function parseOutlineText(text) {
    const lines = text.split(/\r?\n/);
    const parsedBlocks = [];
    
    let currentListBlock = null;
    let currentCodeBlock = null;
    let inCode = false;
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const trimmed = line.trim();
        
        // Handle Code blocks (``` or ~~~)
        if (trimmed.startsWith('```') || trimmed.startsWith('~~~')) {
            if (inCode) {
                // End code block
                if (currentCodeBlock) {
                    parsedBlocks.push(currentCodeBlock);
                    currentCodeBlock = null;
                }
                inCode = false;
            } else {
                // Start code block
                if (currentListBlock) {
                    parsedBlocks.push(currentListBlock);
                    currentListBlock = null;
                }
                inCode = true;
                currentCodeBlock = {
                    id: generateId() + '-outline-' + i,
                    type: 'code',
                    content: ''
                };
            }
            continue;
        }
        
        if (inCode) {
            if (currentCodeBlock) {
                currentCodeBlock.content += (currentCodeBlock.content ? '\n' : '') + line;
            }
            continue;
        }
        
        // Skip empty lines
        if (trimmed === '') {
            if (currentListBlock) {
                parsedBlocks.push(currentListBlock);
                currentListBlock = null;
            }
            continue;
        }
        
        // Handle headings
        if (trimmed.startsWith('#')) {
            if (currentListBlock) {
                parsedBlocks.push(currentListBlock);
                currentListBlock = null;
            }
            
            const match = trimmed.match(/^(#{1,6})\s+(.*)$/);
            if (match) {
                const level = Math.min(3, match[1].length);
                const content = match[2];
                parsedBlocks.push({
                    id: generateId() + '-outline-' + i,
                    type: 'heading',
                    level: level,
                    content: content
                });
                continue;
            }
        }
        
        // Handle lists (- or * or numbered list like 1.)
        const listMatch = trimmed.match(/^[-*+]\s+(.*)$/) || trimmed.match(/^\d+\.\s+(.*)$/);
        if (listMatch) {
            const itemContent = listMatch[1];
            const isNumbered = trimmed.match(/^\d+\./);
            
            if (!currentListBlock) {
                currentListBlock = {
                    id: generateId() + '-outline-' + i,
                    type: 'list',
                    list_style: isNumbered ? 'numbered' : 'bullet',
                    items: []
                };
            }
            currentListBlock.items.push(itemContent);
            continue;
        }
        
        // Flush list block if we hit a paragraph and list is active
        if (currentListBlock) {
            parsedBlocks.push(currentListBlock);
            currentListBlock = null;
        }
        
        // Check if the line is a screenshot placeholder, e.g. [screenshot] or [screenshot: caption]
        const isScreenshotMarker = trimmed.toLowerCase().includes('[screenshot') || 
                                   trimmed.toLowerCase().includes('[image') ||
                                   trimmed.startsWith('![') ||
                                   trimmed.toLowerCase().includes('[paste');
                                   
        if (isScreenshotMarker) {
            let caption = '';
            const mdImgMatch = trimmed.match(/^!\[(.*?)\]/);
            if (mdImgMatch) {
                caption = mdImgMatch[1];
            } else {
                // Try extracting after a colon
                const colonMatch = trimmed.match(/\[screenshot\s*:\s*(.*?)\]/i) || trimmed.match(/\[image\s*:\s*(.*?)\]/i);
                if (colonMatch) {
                    caption = colonMatch[1];
                } else {
                    caption = 'Screenshot Description';
                }
            }
            
            parsedBlocks.push({
                id: generateId() + '-outline-' + i,
                type: 'image',
                image_file: null,
                caption: caption,
                width: 6.0
            });
            continue;
        }
        
        // Default to paragraph block
        parsedBlocks.push({
            id: generateId() + '-outline-' + i,
            type: 'paragraph',
            content: trimmed
        });
    }
    
    // Final flushes
    if (currentListBlock) parsedBlocks.push(currentListBlock);
    if (currentCodeBlock) parsedBlocks.push(currentCodeBlock);
    
    return parsedBlocks;
}

async function generateFromTextTemplate() {
    showLoading(true, 'Compiling text template...');
    
    let filledContent = uploadedTextTemplateContent;
    const imageMap = {};
    
    templateVariables.forEach(variable => {
        const re = new RegExp('\\{\\{\\s*' + variable.name + '\\s*\\}\\}', 'g');
        
        if (variable.type === 'text') {
            const input = document.getElementById(`tpl-field-${variable.name}`);
            const val = input ? input.value : '';
            filledContent = filledContent.replace(re, val);
        } else {
            const fileInput = document.getElementById(`tpl-file-input-${variable.name}`);
            const fileData = fileInput ? fileInput.fileData : null;
            
            const instructionInput = document.getElementById(`tpl-instruction-${variable.name}`);
            const instructionText = instructionInput ? instructionInput.value.trim() : '';
            
            if (fileData) {
                imageMap[variable.name] = fileData;
                let replacement = `[screenshot: ${variable.name}]`;
                if (instructionText) {
                    replacement += `\n${instructionText}`;
                }
                filledContent = filledContent.replace(re, replacement);
            } else {
                filledContent = filledContent.replace(re, '');
            }
        }
    });
    
    // Parse the filled markdown outline into blocks
    const parsedBlocks = parseOutlineText(filledContent);
    
    // Bind uploaded files to the newly parsed image blocks
    parsedBlocks.forEach(block => {
        if (block.type === 'image') {
            const placeholderName = block.caption;
            if (imageMap[placeholderName]) {
                block.image_file = imageMap[placeholderName];
                block.caption = ''; // clear caption placeholder
            }
        }
    });
    
    // Submit to /api/generate-from-scratch
    showLoading(true, 'Generating Word Document...');
    const formData = new FormData();
    
    const docTitle = document.getElementById('doc-title').value.trim() || uploadedTemplateFile.name.replace(/\.[^/.]+$/, "");
    
    const documentData = {
        title: docTitle,
        style: 'standard', // default standard report layout
        blocks: []
    };
    
    parsedBlocks.forEach((block, index) => {
        const item = {
            type: block.type,
            level: block.level,
            content: block.content,
            caption: block.caption,
            width: block.width,
            list_style: block.list_style,
            items: block.items
        };
        
        if (block.type === 'image' && block.image_file) {
            const imageId = `image_${index}`;
            item.image_id = imageId;
            formData.append(imageId, block.image_file);
        }
        
        documentData.blocks.push(item);
    });
    
    formData.append('document_data', JSON.stringify(documentData));
    
    if (uploadedBaseTemplateFile) {
        formData.append('base_template', uploadedBaseTemplateFile);
    }
    
    try {

        const response = await fetch(`${API_BASE}/generate-from-scratch`, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            
            const safeTitle = docTitle.replace(/[^a-z0-9]/gi, '_').toLowerCase();
            a.download = `${safeTitle || 'document'}.docx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            showToast('Document downloaded successfully!', 'success');
        } else {
            const data = await response.json();
            showToast(data.error || 'Failed to generate document', 'error');
        }
    } catch (error) {
        showToast('Error connecting to backend server', 'error');
    } finally {
        showLoading(false);
    }
}

// Base Template Upload Handlers
function handleBaseTemplateUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    if (!file.name.endsWith('.docx')) {
        showToast('Please select a valid Word (.docx) document!', 'error');
        return;
    }
    
    uploadedBaseTemplateFile = file;
    document.getElementById('base-template-name').textContent = file.name;
    document.getElementById('base-template-info').style.display = 'flex';
    showToast('Base styling template loaded!', 'success');
}

function clearBaseTemplate(e) {
    if (e) {
        e.stopPropagation();
        e.preventDefault();
    }
    uploadedBaseTemplateFile = null;
    document.getElementById('base-template-info').style.display = 'none';
    document.getElementById('base-template-file').value = '';
    showToast('Base template removed', 'info');
}



