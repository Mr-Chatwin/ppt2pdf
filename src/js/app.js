(function () {
    'use strict';

    if (!window.__TAURI__) {
        console.error('未在 Tauri 环境中运行');
        document.querySelector('.subtitle').textContent = '运行环境错误，请使用 Tauri 启动';
        return;
    }

    const { invoke } = window.__TAURI__.core;

    // DOM Elements
    const setupSection = document.getElementById('setupSection');
    const dropZone = document.getElementById('dropZone');
    const resultsSection = document.getElementById('resultsSection');
    const fileList = document.getElementById('fileList');
    const fileInputLabel = document.getElementById('fileInputLabel');
    const tabOffice2Pdf = document.getElementById('tabOffice2Pdf');
    const tabPdf2Image = document.getElementById('tabPdf2Image');
    const settingsPanel = document.getElementById('settingsPanel');

    // 状态管理
    let conversionQueue = [];
    let isConverting = false;
    let stopRequested = false;
    let currentMode = 'office2pdf';
    let hasOfficeEngine = false;

    // --- 初始化 ---
    async function init() {
        const subtitle = document.querySelector('.subtitle');
        try {
            const status = await invoke('check_engines');
            // status -> { available: true, engine: 'office', all_engines: ['office', 'wps'] }
            
            if (status.available) {
                hasOfficeEngine = true;
                const engineNames = { office: 'Microsoft Office', wps: 'WPS Office', libreoffice: 'LibreOffice' };
                const engineName = engineNames[status.engine] || status.engine;
                showToast(`已检测到 ${engineName} 转换引擎`, 'success');

                if (status.all_engines && status.all_engines.length > 1) {
                    const options = status.all_engines.map(e => 
                        `<option value="${e}" ${e === status.engine ? 'selected' : ''}>${engineNames[e] || e}</option>`
                    ).join('');
                    
                    subtitle.innerHTML = `转换引擎：<select id="engineSelect" class="engine-select">${options}</select> · 高质量 · 极速`;
                    
                    document.getElementById('engineSelect').addEventListener('change', async (e) => {
                        const next = e.target.value;
                        const ok = await invoke('set_engine', { engine: next });
                        if (ok) {
                            showToast(`已切换到 ${engineNames[next] || next} 引擎`, 'info');
                        } else {
                            showToast(`切换引擎失败`, 'error');
                            e.target.value = status.engine;
                        }
                    });
                } else {
                    subtitle.textContent = `使用 ${engineName} 引擎 · 高质量转换 · 极速`;
                }

                dropZone.classList.remove('hidden');
            } else {
                hasOfficeEngine = false;
                subtitle.textContent = '缺少转换引擎';
                setupSection.classList.remove('hidden');
            }
        } catch (e) {
            hasOfficeEngine = false;
            subtitle.textContent = '引擎检测失败';
            showToast('引擎检测出错: ' + e, 'error');
            dropZone.classList.remove('hidden');
        }

        bindEvents();
    }

    function bindEvents() {
        // Mode Tabs Click
        tabOffice2Pdf.addEventListener('click', () => switchMode('office2pdf'));
        tabPdf2Image.addEventListener('click', () => switchMode('pdf2image'));

        // 使用后台 rfd 原生弹窗
        fileInputLabel.addEventListener('click', async () => {
            if (isConverting) return;
            try {
                const paths = await invoke('select_files', { mode: currentMode });
                if (paths && paths.length > 0) {
                    addFilesToQueue(paths);
                }
            } catch (err) {
                console.error('选择文件出错: ', err);
            }
        });

        // 强行阻止 HTML 默认的拖拽拦截，让 Tauri 事件生效
        document.addEventListener('dragover', e => e.preventDefault());
        document.addEventListener('drop', e => e.preventDefault());

        // 监听 Tauri 提供的拖拽事件
        window.__TAURI__.event.listen('tauri://drag-drop', (event) => {
            if (isConverting) return;
            const paths = event.payload.paths || event.payload; // 兼容不同版本的 payload 结构
            if (!Array.isArray(paths)) return;
            
            if (currentMode === 'office2pdf') {
                const validPaths = paths.filter(p => /\.(pptx?|ppsx?|pps|docx?|doc|xlsx?|xls|csv)$/i.test(p));
                if (validPaths.length > 0) {
                    addFilesToQueue(validPaths);
                } else {
                    showToast('请拖入有效的 Office(Word/Excel/PPT) 文件', 'error');
                }
            } else {
                const validPaths = paths.filter(p => /\.pdf$/i.test(p));
                if (validPaths.length > 0) {
                    addFilesToQueue(validPaths);
                } else {
                    showToast('请拖入有效的 PDF 文件', 'error');
                }
            }
            dropZone.classList.remove('drag-over');
        });
        
        window.__TAURI__.event.listen('tauri://drag-enter', () => dropZone.classList.add('drag-over'));
        window.__TAURI__.event.listen('tauri://drag-leave', () => dropZone.classList.remove('drag-over'));

        document.getElementById('startConvertBtn')?.addEventListener('click', startConversion);
        document.getElementById('stopConvertBtn')?.addEventListener('click', () => { stopRequested = true; });
        document.getElementById('newConvertBtn')?.addEventListener('click', clearQueue);
        document.getElementById('addMoreBtn')?.addEventListener('click', () => fileInputLabel.click());
    }

    function switchMode(mode) {
        if (isConverting) {
            showToast('正在转换中，无法切换模式', 'warning');
            return;
        }
        if (conversionQueue.length > 0) {
            if (!confirm('切换模式将清空当前文件列表，是否继续？')) {
                return;
            }
            clearQueue();
        }
        currentMode = mode;

        if (mode === 'office2pdf') {
            tabOffice2Pdf.classList.add('active');
            tabPdf2Image.classList.remove('active');
            settingsPanel.classList.add('hidden');

            document.querySelector('#dropZone h2').textContent = '拖拽 Office 文件到这里';
            document.querySelector('.drop-zone-hint').textContent = '支持 .docx / .xlsx / .pptx 以及对应的老版本格式格式，可同时拖放';

            document.querySelector('.subtitle').classList.remove('hidden');

            if (hasOfficeEngine) {
                setupSection.classList.add('hidden');
                dropZone.classList.remove('hidden');
            } else {
                setupSection.classList.remove('hidden');
                dropZone.classList.add('hidden');
            }
        } else {
            tabOffice2Pdf.classList.remove('active');
            tabPdf2Image.classList.add('active');
            settingsPanel.classList.remove('hidden');

            document.querySelector('#dropZone h2').textContent = '拖拽 PDF 文件到这里';
            document.querySelector('.drop-zone-hint').textContent = '支持 .pdf 格式，可同时拖放';

            document.querySelector('.subtitle').classList.add('hidden');

            setupSection.classList.add('hidden');
            dropZone.classList.remove('hidden');
        }
    }

    function addFilesToQueue(paths) {
        dropZone.classList.add('hidden');
        resultsSection.classList.remove('hidden');

        paths.forEach(p => {
            const name = p.split(/[\\/]/).pop();
            conversionQueue.push({ path: p, name, status: 'pending', error: null, result: null });
        });
        renderQueue();
    }

    function clearQueue() {
        if (isConverting) {
            showToast('请先停止当前转换', 'error');
            return;
        }
        conversionQueue = [];
        resultsSection.classList.add('hidden');
        dropZone.classList.remove('hidden');
    }

    // 供 HTML 内联 onClick 调用的全局函数
    window.removeQueueItem = function(index) {
        if (isConverting) return;
        conversionQueue.splice(index, 1);
        if (conversionQueue.length === 0) clearQueue();
        else renderQueue();
    };

    window.openFile = async function(path) {
        try {
            await invoke('open_file', { path });
        } catch (e) {
            showToast('打开文件失败', 'error');
        }
    };

    window.openFolder = async function(path) {
        try {
            await invoke('open_folder', { path });
        } catch (e) {
            showToast('打开文件夹失败', 'error');
        }
    };

    // --- 界面渲染逻辑 ---
    function renderQueue() {
        const resultsTitle = document.getElementById('resultsTitle');
        resultsTitle.textContent = `待转换 ${conversionQueue.length} 个文件`;
        fileList.innerHTML = '';

        conversionQueue.forEach((item, i) => {
            const div = document.createElement('div');
            div.className = 'file-item';
            
            let statusHtml = '<span style="color:var(--text-secondary)">⏳ 等待中</span>';
            let statusClass = '';
            
            if (item.status === 'converting') {
                statusHtml = '<span style="color:var(--accent-1)">🔄 转换中...</span>';
                statusClass = 'converting';
            } else if (item.status === 'success') {
                statusHtml = '<span style="color:var(--success)">✅ 完成</span>';
                statusClass = 'success';
            } else if (item.status === 'error') {
                statusHtml = `<span style="color:var(--error)">❌ 失败: ${item.error}</span>`;
                statusClass = 'error';
            } else if (item.status === 'stopped') {
                statusHtml = '<span style="color:var(--text-secondary)">⏹️ 已取消</span>';
            }

            let actionsHtml = '';
            if (item.status === 'success' && item.result) {
                if (/\.pdf$/i.test(item.name)) {
                    actionsHtml = `
                        <button class="btn btn-sm btn-ghost" onclick="window.openFolder('${escapeJs(item.result)}')">打开文件夹</button>
                    `;
                } else {
                    actionsHtml = `
                        <button class="btn btn-sm btn-ghost" onclick="window.openFile('${escapeJs(item.result)}')">打开 PDF</button>
                        <button class="btn btn-sm btn-ghost" onclick="window.openFolder('${escapeJs(item.result)}')">打开文件夹</button>
                    `;
                }
            } else if (item.status === 'pending') {
                actionsHtml = `<button class="btn btn-sm btn-ghost btn-danger" onclick="removeQueueItem(${i})">移除</button>`;
            }

            const isPdf = /\.pdf$/i.test(item.name);
            const isWord = /\.(docx?|doc)$/i.test(item.name);
            const isExcel = /\.(xlsx?|xls|csv)$/i.test(item.name);
            const icon = isPdf ? '📄' : (isWord ? '📝' : (isExcel ? '📗' : '📊'));

            div.innerHTML = `
                <div class="file-item-icon">${icon}</div>
                <div class="file-item-info">
                    <div class="file-item-name">${escapeHtml(item.name)}</div>
                    <div class="file-item-status ${statusClass}">${statusHtml}</div>
                </div>
                <div class="file-item-actions">${actionsHtml}</div>
            `;
            fileList.appendChild(div);
        });

        updateConvertButtons();
    }

    function updateConvertButtons() {
        const startConvertBtn = document.getElementById('startConvertBtn');
        const stopConvertBtn = document.getElementById('stopConvertBtn');
        const newConvertBtn = document.getElementById('newConvertBtn');
        const addMoreBtn = document.getElementById('addMoreBtn');
        
        const hasPending = conversionQueue.some(i => i.status === 'pending');
        
        if (isConverting) {
            startConvertBtn.classList.add('hidden');
            stopConvertBtn.classList.remove('hidden');
            newConvertBtn.disabled = true;
            addMoreBtn.disabled = true;
        } else {
            stopConvertBtn.classList.add('hidden');
            startConvertBtn.classList.remove('hidden');
            newConvertBtn.disabled = false;
            addMoreBtn.disabled = false;
            
            if (!hasPending && conversionQueue.length > 0) {
                startConvertBtn.disabled = true;
                startConvertBtn.textContent = '转换已完成';
            } else {
                startConvertBtn.disabled = conversionQueue.length === 0;
                startConvertBtn.innerHTML = `
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>
                    开始转换
                `;
            }
        }
    }

    async function startConversion() {
        if (isConverting) return;
        
        const pendingItems = conversionQueue.filter(i => i.status === 'pending');
        if (pendingItems.length === 0) return;

        isConverting = true;
        stopRequested = false;
        
        const engineSelect = document.getElementById('engineSelect');
        if (engineSelect) engineSelect.disabled = true;
        
        updateConvertButtons();

        // 将要处理的全部标记为状态转换中
        pendingItems.forEach(item => { item.status = 'converting'; });
        renderQueue();

        let successCount = 0;
        let failCount = 0;
        const paths = pendingItems.map(item => item.path);

        try {
            // 根据转换模式分别调用对应的底层 API 引擎
            let results = [];
            if (currentMode === 'office2pdf') {
                results = await invoke('convert_batch', { paths });
            } else {
                const format = document.getElementById('formatSelect').value;
                const dpi = parseFloat(document.getElementById('dpiSelect').value);
                results = await invoke('convert_pdf_to_images', { paths, format, dpi });
            }
            
            results.forEach(res => {
                const item = conversionQueue.find(i => i.path === res.path);
                if (item) {
                    if (res.success) {
                        item.status = 'success';
                        item.result = res.output_path || res.pdf_path;
                        successCount++;
                    } else {
                        item.status = 'error';
                        item.error = res.error_msg || '未知错误';
                        failCount++;
                    }
                }
            });
        } catch (err) {
            console.error(err);
            const errMsg = typeof err === 'string' ? err : (err.message || JSON.stringify(err));
            pendingItems.forEach(item => {
                if (item.status === 'converting') {
                    item.status = 'error';
                    item.error = errMsg;
                    failCount++;
                }
            });
        }

        isConverting = false;
        if (engineSelect) engineSelect.disabled = false;
        updateConvertButtons();
        renderQueue();

        const resultsTitle = document.getElementById('resultsTitle');
        resultsTitle.textContent = `转换完成 — ${successCount} 个成功${failCount > 0 ? `，${failCount} 个失败` : ''}`;
        showToast(`完成！${successCount} 个文件通过批处理引擎完成极速转换`, successCount > 0 ? 'success' : 'error');
    }

    // --- 工具函数 ---
    function showToast(msg, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = msg;
        document.body.appendChild(toast);
        
        toast.offsetHeight; // reflow
        toast.classList.add('show');
        
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 400);
        }, 3000);
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str).replace(/[&<>'"]/g, tag => ({
            '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
        }[tag] || tag));
    }
    
    function escapeJs(str) {
        if (!str) return '';
        return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '\\"');
    }

    init();
})();
