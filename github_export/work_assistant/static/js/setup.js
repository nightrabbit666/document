document.addEventListener('DOMContentLoaded', () => {
    // State
    const uploadedFiles = {
        A: null, // {id, name}
        B: null,
        C: null
    };

    let analyzedParams = [];

    // --- File Upload Logic ---
    
    function setupDropZone(zoneKey, fileInputId) {
        const zone = document.getElementById(`zone${zoneKey}`);
        const input = document.getElementById(fileInputId);
        
        // Drag events
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            zone.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        // Highlight
        ['dragenter', 'dragover'].forEach(eventName => {
            zone.addEventListener(eventName, () => zone.classList.add('bg-slate-100'), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            zone.addEventListener(eventName, () => zone.classList.remove('bg-slate-100'), false);
        });

        // Drop
        zone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            handleFiles(files, zoneKey);
        });

        // Click
        zone.addEventListener('click', () => input.click());
        input.addEventListener('change', function() {
            handleFiles(this.files, zoneKey);
        });
    }

    function handleFiles(files, zoneKey) {
        if (files.length > 0) {
            uploadFile(files[0], zoneKey);
        }
    }

    function uploadFile(file, zoneKey) {
        const formData = new FormData();
        formData.append('file', file);
        
        // Show uploading state (simple)
        const infoDiv = document.getElementById(`fileInfo${zoneKey}`);
        infoDiv.textContent = `上傳中...`;
        infoDiv.classList.remove('hidden');

        fetch('/api/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                uploadedFiles[zoneKey] = {
                    id: data.file_id,
                    name: data.original_name
                };
                
                // Update UI
                infoDiv.textContent = `已就緒: ${data.original_name} (${formatBytes(data.size)})`;
                
                // Hide icon or change style (Optional refinement)
                document.querySelector(`#zone${zoneKey} svg`).classList.remove('text-slate-400'); // if it was gray
                // We keep the icon but labeled file is clear
            } else {
                alert('Upload failed: ' + data.error);
                infoDiv.textContent = '上傳失敗';
            }
        })
        .catch(err => {
            console.error(err);
            alert('Upload error');
        });
    }

    function formatBytes(bytes, decimals = 2) {
        if (!+bytes) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
    }

    // Init Zones
    setupDropZone('A', 'fileA');
    setupDropZone('B', 'fileB');
    setupDropZone('C', 'fileC');

    // --- Step 1 -> 2: Analysis ---

    const btnAnalyze = document.getElementById('btnAnalyze');
    const loadingIcon = document.getElementById('loadingIcon');
    const btnText = document.getElementById('btnText');

    btnAnalyze.addEventListener('click', () => {
        if (!uploadedFiles.A) {
            alert('請至少上傳範本檔案 (Zone A)');
            return;
        }

        // UI Loading
        btnAnalyze.disabled = true;
        loadingIcon.classList.remove('hidden');
        btnText.textContent = '分析文件中...';

        const payload = {
            template_file_id: uploadedFiles.A.id,
            excel_file_id: uploadedFiles.C?.id,
            old_doc_file_id: uploadedFiles.B?.id
        };

        fetch('/api/analyze', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        })
        .then(r => r.json())
        .then(data => {
            if (data.error) {
                alert('Analysis error: ' + data.error);
                resetAnalyzeBtn();
                return;
            }
            
            analyzedParams = data.parameters || [];
            // Updated: Pass logic summary and token usage
            renderParams(analyzedParams, data.diff_report, data.logic_summary, data.token_usage);
            
            // Go to Step 2
            goToStep(2);
        })
        .catch(err => {
            console.error(err);
            alert('Analysis request failed');
            resetAnalyzeBtn();
        });
    });

    function resetAnalyzeBtn() {
        btnAnalyze.disabled = false;
        loadingIcon.classList.add('hidden');
        btnText.textContent = '開始 AI 分析';
    }

    function renderParams(params, diffReport, logicSummary, tokenUsage) {
        const container = document.getElementById('paramsContainer');
        container.innerHTML = '';

        // [Logic Summary Section]
        if (logicSummary) {
             const logicBox = document.createElement('div');
             logicBox.className = 'mb-6 bg-emerald-50 border border-emerald-200 p-4 rounded text-sm text-emerald-800';
             logicBox.innerHTML = `
                <h4 class="font-bold flex items-center mb-2">
                    <svg class="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                    專案邏輯分析 (AI Generated)
                </h4>
                <p class="whitespace-pre-wrap">${logicSummary}</p>
             `;
             container.appendChild(logicBox);
        }
        
        // [Token Usage Section]
        if (tokenUsage) {
             const tokenBox = document.createElement('div');
             // Small discrete style
             tokenBox.className = 'mb-2 flex justify-end text-xs text-slate-400 gap-4';
             tokenBox.innerHTML = `
                <span>Prompt: ${tokenUsage.prompt_tokens}</span>
                <span>Candidates: ${tokenUsage.candidates_tokens}</span>
                <span class="font-bold">Total: ${tokenUsage.total_tokens} Tokens</span>
             `;
             // Insert before logic box or params
             container.appendChild(tokenBox);
        }

        // [Diff Report Section]
        if (diffReport) {
            const reportBox = document.createElement('div');
            reportBox.className = 'mb-6 bg-slate-800 text-slate-200 p-4 rounded text-xs font-mono overflow-auto max-h-60 whitespace-pre';
            reportBox.textContent = `=== 結構比對報告 (Cross-Validation) ===\n${diffReport}`;
            container.appendChild(reportBox);
        }

        if (!params || params.length === 0) {
            const emptyMsg = document.createElement('div');
            emptyMsg.className = "text-center py-10 text-slate-500";
            emptyMsg.innerHTML = `<p class="text-lg">⚠️ 未偵測到變數</p><p class="text-sm">請確認範本與參考文件是否有顯著差異。</p>`;
            container.appendChild(emptyMsg);
            return;
        }

        params.forEach((p, idx) => {
            const card = document.createElement('div');
            card.className = 'border border-slate-200 rounded p-4 flex flex-col md:flex-row gap-4 items-start';
            
            card.innerHTML = `
                <div class="flex-1 w-full">
                    <div class="flex justify-between mb-2">
                        <label class="font-bold text-slate-700">變數名稱: 
                            <input type="text" data-idx="${idx}" class="param-name border-b border-dotted border-slate-400 focus:outline-none focus:border-emerald-500 bg-transparent" value="${p.name}">
                        </label>
                        <span class="text-xs bg-slate-100 text-slate-500 px-2 py-1 rounded">${p.type}</span>
                    </div>
                    <div class="mb-2">
                         <label class="block text-xs text-slate-500">描述</label>
                         <input type="text" data-idx="${idx}" class="param-desc w-full border border-slate-200 rounded px-2 py-1 text-sm text-slate-700" value="${p.description}">
                    </div>
                    <div class="bg-slate-50 p-2 rounded text-xs text-slate-500 italic">
                        上下文: "${p.context || '...'}"
                    </div>
                </div>
                <!-- Example Value Preview -->
                <div class="md:w-1/3 w-full bg-blue-50 p-3 rounded border border-blue-100">
                    <p class="text-xs text-blue-500 font-bold mb-1">參考值 (Example)</p>
                    <p class="text-sm text-blue-800 break-all">${p.example || '無'}</p>
                </div>
            `;
            container.appendChild(card);
        });

        // Bind inputs to array
        container.querySelectorAll('input.param-name').forEach(inp => {
            inp.addEventListener('change', (e) => {
                const idx = e.target.getAttribute('data-idx');
                analyzedParams[idx].name = e.target.value;
            });
        });

        container.querySelectorAll('input.param-desc').forEach(inp => {
            inp.addEventListener('change', (e) => {
                const idx = e.target.getAttribute('data-idx');
                analyzedParams[idx].description = e.target.value;
            });
        });
    }

    // --- Step Navigation ---
    function goToStep(step) {
        // Hide all
        document.getElementById('step-1').classList.add('hidden');
        document.getElementById('step-2').classList.add('hidden');
        document.getElementById('step-3').classList.add('hidden');
        
        // Show target
        document.getElementById(`step-${step}`).classList.remove('hidden');

        // Update Indicators
        for (let i = 1; i <= 3; i++) {
            const ind = document.getElementById(`step-ind-${i}`);
            if (i === step) {
                ind.classList.remove('text-slate-400');
                ind.classList.add('text-emerald-600');
                const circ = ind.querySelector('span');
                circ.classList.remove('bg-slate-200', 'text-slate-500');
                circ.classList.add('bg-emerald-600', 'text-white');
            } else if (i < step) {
                // Completed
                ind.classList.remove('text-slate-400');
                ind.classList.add('text-emerald-600');
                const circ = ind.querySelector('span');
                circ.classList.remove('bg-slate-200', 'text-slate-500');
                circ.classList.add('bg-emerald-600', 'text-white');
            }
        }
    }

    // --- Step 2 -> 3: Save ---
    
    // Mode Toggles
    const modeOneShot = document.getElementById('modeOneShot');
    const modeMonthly = document.getElementById('modeMonthly');
    const projectModeInput = document.getElementById('projectMode');

    if (modeOneShot && modeMonthly) {
        modeOneShot.addEventListener('click', () => {
            setMode('one_shot');
        });
        modeMonthly.addEventListener('click', () => {
            setMode('monthly');
        });
    }

    function setMode(mode) {
        projectModeInput.value = mode;
        if (mode === 'one_shot') {
            modeOneShot.className = 'flex-1 py-1 px-2 text-xs border rounded bg-emerald-100 text-emerald-700 border-emerald-300 font-bold';
            modeMonthly.className = 'flex-1 py-1 px-2 text-xs border rounded bg-white text-slate-500 border-slate-200 hover:bg-slate-50';
        } else {
            modeOneShot.className = 'flex-1 py-1 px-2 text-xs border rounded bg-white text-slate-500 border-slate-200 hover:bg-slate-50';
            modeMonthly.className = 'flex-1 py-1 px-2 text-xs border rounded bg-blue-100 text-blue-700 border-blue-300 font-bold';
        }
    }
    
    document.getElementById('btnBack').addEventListener('click', () => {
        goToStep(1);
        resetAnalyzeBtn();
    });

    document.getElementById('btnSave').addEventListener('click', () => {
        const name = document.getElementById('projectName').value;
        if (!name) {
            alert('請輸入專案名稱');
            return;
        }

        const payload = {
            project_name: name,
            project_desc: document.getElementById('projectDesc').value,
            mode: projectModeInput ? projectModeInput.value : 'one_shot',
            features: {
                daily: document.getElementById('featureDaily')?.checked || false,
                monthly: document.getElementById('featureMonthly')?.checked || false,
                debug: document.getElementById('featureDebug')?.checked || false
            },
            template_file_id: uploadedFiles.A.id,
            parameters: analyzedParams
        };

        fetch('/api/save_project', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        })
        .then(r => r.json())
        .then(data => {
            if (data.success) {
                document.getElementById('linkStart').href = `/project/${data.project_id}`;
                goToStep(3);
            } else {
                alert('Save failed');
            }
        });
    });

});
