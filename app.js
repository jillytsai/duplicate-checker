document.addEventListener('DOMContentLoaded', () => {
    const fileUpload = document.getElementById('file-upload');
    const dropzone = document.getElementById('dropzone');
    const processingPanel = document.getElementById('processing-panel');
    const resultsPanel = document.getElementById('results-panel');
    const resultsTbody = document.getElementById('results-tbody');
    const processStatus = document.getElementById('process-status');
    const processCount = document.getElementById('process-count');
    const progressBar = document.getElementById('progress-bar');
    const downloadBtn = document.getElementById('download-btn');
    const resetBtn = document.getElementById('reset-btn');
    const stopBtn = document.getElementById('stop-btn');

    let originalWorkbook = null;
    let targetSheetName = null;
    let excelData = [];
    let processedData = [];
    let isbnKey = null;
    let titleKey = null;
    let authorKey = null;
    let publisherKey = null;
    let isProcessing = false;

    // Drag and drop setup
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => dropzone.classList.add('dragover'));
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => dropzone.classList.remove('dragover'));
    });

    dropzone.addEventListener('drop', (e) => {
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });

    fileUpload.addEventListener('change', (e) => {
        if (e.target.files.length) handleFile(e.target.files[0]);
    });

    // Button interactions
    resetBtn.addEventListener('click', () => location.reload());
    downloadBtn.addEventListener('click', () => exportExcel());

    stopBtn.addEventListener('click', () => {
        isProcessing = false;
        stopBtn.disabled = true;
        stopBtn.innerHTML = '<i class="ph-bold ph-spinner ph-spin"></i> 正在停止...';
    });

    // File processing
    function handleFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                originalWorkbook = workbook;
                originalWorkbook = workbook;
                excelData = [];
                let skippedSheets = [];
                
                for (let i = 0; i < workbook.SheetNames.length; i++) {
                    const sheetName = workbook.SheetNames[i];
                    
                    // 檢查分頁是否為隱藏狀態 (Hidden: 1 = 隱藏, 2 = 深度隱藏)
                    if (workbook.Workbook && workbook.Workbook.Sheets && workbook.Workbook.Sheets[i]) {
                        const hiddenState = workbook.Workbook.Sheets[i].Hidden;
                        if (hiddenState === 1 || hiddenState === 2) {
                            console.log(`跳過隱藏的分頁: ${sheetName}`);
                            continue;
                        }
                    }

                    const worksheet = workbook.Sheets[sheetName];
                    let rawRows;
                    if (worksheet['!ref']) {
                        const range = XLSX.utils.decode_range(worksheet['!ref']);
                        range.s.c = 0; // 強制從 A 欄開始，防止前面有空欄導致 index 位移而讀錯出版年
                        rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", range: XLSX.utils.encode_range(range) });
                    } else {
                        rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                    }
                    
                    if (rawRows.length === 0) continue;

                    let headerRowIndex = -1;
                    let isbnColIndex = -1;
                    let eissnColIndex = -1;
                    let eisbnColIndex = -1;
                    let pisbnColIndex = -1;
                    let titleColIndex = -1;
                    let authorColIndex = -1;
                    let publisherColIndex = -1;
                    let pubYearColIndex = -1;
                    let seqColIndex = -1;

                    // 尋找包含標題的列 (Header Row)
                    for (let r = 0; r < rawRows.length; r++) {
                        const row = rawRows[r];
                        if (!Array.isArray(row)) continue;

                        let hasIsbn = false;
                        let hasTitle = false;

                        for (let c = 0; c < row.length; c++) {
                            let cellStr = String(row[c]).replace(/\s+/g, '').toLowerCase();
                            if (eissnColIndex === -1 && (cellStr.includes('eissn') || cellStr === 'e-issn' || cellStr === 'onlineissn')) {
                                hasIsbn = true;
                                eissnColIndex = c;
                            } else if (eisbnColIndex === -1 && (cellStr.includes('eisbn13') || cellStr.includes('eisbn'))) {
                                hasIsbn = true;
                                eisbnColIndex = c;
                            } else if (pisbnColIndex === -1 && (cellStr.includes('pisbn13') || cellStr.includes('pisbn'))) {
                                hasIsbn = true;
                                pisbnColIndex = c;
                            } else if (isbnColIndex === -1 && (cellStr.includes('isbn') || cellStr.includes('issn') || cellStr.includes('條碼') || cellStr.includes('barcode'))) {
                                hasIsbn = true;
                                isbnColIndex = c;
                            }
                            if (titleColIndex === -1 && (cellStr.includes('書名') || cellStr.includes('刊名') || cellStr.includes('題名') || cellStr.includes('title') || cellStr.includes('bookname') || cellStr.includes('booktitle') || cellStr.includes('name') || cellStr.includes('subject'))) {
                                hasTitle = true;
                                titleColIndex = c;
                            }
                            if (authorColIndex === -1 && (cellStr.includes('作者') || cellStr.includes('著者') || cellStr.includes('author') || cellStr.includes('creator') || cellStr.includes('writer'))) {
                                authorColIndex = c;
                            }
                            // 嚴格比對出版者、出版社
                            if (publisherColIndex === -1 && (cellStr.includes('出版者') || cellStr.includes('出版社') || cellStr === '出版' || cellStr.includes('publisher') || cellStr.includes('press') || cellStr.includes('publishing') || cellStr === 'pub')) {
                                publisherColIndex = c;
                            }
                            // 獨立比對出版年
                            if (pubYearColIndex === -1 && (cellStr.includes('出版年') || cellStr.includes('pubyear') || cellStr.includes('year') || cellStr.includes('date') || cellStr.includes('pubdate') || cellStr.includes('publishyear'))) {
                                pubYearColIndex = c;
                            }
                            if (seqColIndex === -1 && (cellStr.includes('序號') || cellStr === 'no' || cellStr === 'id' || cellStr === 'seq' || cellStr === 'index' || cellStr === 'number')) {
                                seqColIndex = c;
                            }
                        }

                        if (hasIsbn || hasTitle) {
                            headerRowIndex = r;
                            break;
                        }
                    }

                    if (headerRowIndex !== -1) {
                        const headerRow = rawRows[headerRowIndex];
                        for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
                            const rowArr = rawRows[r];
                            
                            const isRowEmpty = rowArr.every(cell => String(cell).trim() === '');
                            if (isRowEmpty) continue;

                            let rowObj = {};
                            // 用找到的標題列當作物件的 Key，才能保留匯出時的正確標題
                            for (let c = 0; c < headerRow.length; c++) {
                                let hdr = String(headerRow[c]).trim() || `__EMPTY_${c}`;
                                let cellVal = rowArr[c] !== undefined ? rowArr[c] : "";
                                
                                // 特別處理出版年：若是 Excel 日期序號，將其轉換為完整的 YYYY/M/D 格式，避免年份只顯示兩碼 (例如 26)
                                if (c === pubYearColIndex) {
                                    const cellRef = XLSX.utils.encode_cell({ c: c, r: r });
                                    const cell = worksheet[cellRef];
                                    if (cell && cell.v !== undefined) {
                                        // 判斷是否為 Excel 序號日期 (大於 10000 且顯示文字帶有日期分隔符號)
                                        if (typeof cell.v === 'number' && cell.v > 10000 && cell.w && (cell.w.includes('/') || cell.w.includes('-') || cell.w.includes('年'))) {
                                            const date = new Date(Math.round((cell.v - 25569) * 86400 * 1000));
                                            const year = date.getUTCFullYear();
                                            const month = date.getUTCMonth() + 1;
                                            const day = date.getUTCDate();
                                            cellVal = `${year}/${month}/${day}`;
                                        } else if (cell.w !== undefined) {
                                            cellVal = String(cell.w);
                                        } else {
                                            cellVal = String(cell.v);
                                        }
                                    }
                                }
                                
                                rowObj[hdr] = cellVal;
                            }
                            
                            rowObj['__sheetName'] = sheetName;
                            rowObj['__SYS_ISBN'] = isbnColIndex !== -1 ? String(rowArr[isbnColIndex] !== undefined ? rowArr[isbnColIndex] : "") : "";
                            rowObj['__SYS_EISSN'] = eissnColIndex !== -1 ? String(rowArr[eissnColIndex] !== undefined ? rowArr[eissnColIndex] : "") : "";
                            rowObj['__SYS_EISBN'] = eisbnColIndex !== -1 ? String(rowArr[eisbnColIndex] !== undefined ? rowArr[eisbnColIndex] : "") : "";
                            rowObj['__SYS_PISBN'] = pisbnColIndex !== -1 ? String(rowArr[pisbnColIndex] !== undefined ? rowArr[pisbnColIndex] : "") : "";
                            rowObj['__SYS_TITLE'] = titleColIndex !== -1 ? String(rowArr[titleColIndex] !== undefined ? rowArr[titleColIndex] : "") : "";
                            rowObj['__SYS_AUTHOR'] = authorColIndex !== -1 ? String(rowArr[authorColIndex] !== undefined ? rowArr[authorColIndex] : "") : "";
                            rowObj['__SYS_PUBLISHER'] = publisherColIndex !== -1 ? String(rowArr[publisherColIndex] !== undefined ? rowArr[publisherColIndex] : "") : "";
                            
                            let pubYearVal = "";
                            if (pubYearColIndex !== -1) {
                                let hdr = String(headerRow[pubYearColIndex]).trim() || `__EMPTY_${pubYearColIndex}`;
                                pubYearVal = String(rowObj[hdr] !== undefined ? rowObj[hdr] : "");
                            }
                            rowObj['__SYS_PUBYEAR'] = pubYearVal;
                            
                            if (seqColIndex !== -1) {
                                let hdr = String(headerRow[seqColIndex]).trim() || `__EMPTY_${seqColIndex}`;
                                rowObj['__SYS_SEQ_KEY'] = hdr;
                            }

                            // 只有包含 ISBN/ISSN/EISSN/eISBN/pISBN 或是 書名/刊名 的列才推進佇列，防止讀到最後面的合計、備註欄等等
                            if (rowObj['__SYS_ISBN'] || rowObj['__SYS_EISSN'] || rowObj['__SYS_EISBN'] || rowObj['__SYS_PISBN'] || rowObj['__SYS_TITLE']) {
                                excelData.push(rowObj);
                            }
                        }
                    } else {
                        console.warn("分頁: " + sheetName + " 找不到包含 ISBN/ISSN/EISSN 或 書名/刊名 的標題列！");
                        skippedSheets.push(sheetName);
                    }
                }
                
                const totalSheets = workbook.SheetNames.length;
                console.log("發現分頁數量:", totalSheets, "分頁名稱:", workbook.SheetNames);
                console.log("讀取到的總資料筆數:", excelData.length);
                
                if (skippedSheets.length > 0) {
                    alert(`提醒：以下 ${skippedSheets.length} 個分頁因為找不到「ISBN / ISSN / EISSN」或「書名 / 刊名」標題列，已被自動略過：\n👉 ${skippedSheets.join(", ")}\n(若這些分頁內真的沒有書籍/期刊資料，請忽略此訊息)`);
                }

                if (excelData.length === 0) {
                    alert("檔案為空！或是所有分頁都找不到為「ISBN / ISSN / EISSN」或「書名 / 刊名」的欄位列。");
                    return;
                }

                startProcessing();
            } catch (err) {
                console.error(err);
                alert("讀取 Excel 失敗：" + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // Checking flow
    async function startProcessing() {
        dropzone.style.display = 'none';
        processingPanel.style.display = 'block';
        resultsPanel.style.display = 'block';
        stopBtn.style.display = 'inline-flex';
        stopBtn.disabled = false;
        stopBtn.innerHTML = '<i class="ph-bold ph-stop-circle"></i> 停止查詢';
        isProcessing = true;
        
        processedData = [];

        for (let i = 0; i < excelData.length; i++) {
            if (!isProcessing) {
                processStatus.textContent = `已手動停止！您可以隨時匯出已處理的部分。`;
                break;
            }

            const row = { ...excelData[i] };
            
            // 自動校正原本 Excel 中的「序號」欄位，確保從 1 開始排列
            if (row['__SYS_SEQ_KEY']) {
                row[row['__SYS_SEQ_KEY']] = i + 1;
            }

            let rawIsbn = row['__SYS_ISBN'] ? row['__SYS_ISBN'].replace(/[- ]/g, '').trim() : '';
            // Use only alphanumeric prefixes for ISBNs to ignore attached metadata 
            let valIsbn = rawIsbn.split(/[^0-9X]/i)[0]; 
            
            let rawEisbn = row['__SYS_EISBN'] ? row['__SYS_EISBN'].replace(/[- ]/g, '').trim() : '';
            let valEisbn = rawEisbn.split(/[^0-9X]/i)[0];

            let rawPisbn = row['__SYS_PISBN'] ? row['__SYS_PISBN'].replace(/[- ]/g, '').trim() : '';
            let valPisbn = rawPisbn.split(/[^0-9X]/i)[0];

            let rawEissn = row['__SYS_EISSN'] ? row['__SYS_EISSN'].replace(/[- ]/g, '').trim() : '';
            let valEissn = rawEissn.split(/[^0-9X]/i)[0]; 

            let title = row['__SYS_TITLE'] ? row['__SYS_TITLE'].trim() : '';
            let author = row['__SYS_AUTHOR'].trim();
            let publisher = row['__SYS_PUBLISHER'].trim();
            let pubYear = row['__SYS_PUBYEAR'].trim();

            let barcodes = [];
            if (valIsbn) barcodes.push(valIsbn);
            if (valEisbn) barcodes.push(valEisbn);
            if (valPisbn) barcodes.push(valPisbn);
            if (valEissn) barcodes.push(valEissn);
            let displayBarcode = barcodes.length > 0 ? barcodes.join(" / ") : "-";

            processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): ${title || displayBarcode || '未知'}`;
            processCount.textContent = `${i + 1} / ${excelData.length}`;
            progressBar.style.width = `${((i + 1) / excelData.length) * 100}%`;
            
            let rowStatus = "fetching";
            let statusText = "查詢中...";
            let isDuplicate = "否"; 
            let specialProperty = "查詢中...";
            let collectionQuantity = "查詢中...";
            let nptuUrl = "";
            
            createOrUpdateRow(i, displayBarcode, title, author, publisher, pubYear, specialProperty, rowStatus, statusText, collectionQuantity, nptuUrl);

            try {
                if (i > 0) await new Promise(r => setTimeout(r, 600)); 
                
                let foundMatch = null;
                
                // 1. Try searching by ISBN / ISSN
                if (valIsbn && valIsbn.length >= 8) {
                    let searchType = valIsbn.length === 8 ? 'ISSN' : 'ISBN';
                    processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): [查${searchType}] ${valIsbn}`;
                    const res = await checkNptu(valIsbn, 'k');
                    if (res.found) { foundMatch = searchType; collectionQuantity = res.quantity; nptuUrl = res.url; }
                }
                
                // 1.2 Try searching by eISBN
                if (!foundMatch && valEisbn && valEisbn.length >= 8) {
                    processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): [查eISBN] ${valEisbn}`;
                    await new Promise(r => setTimeout(r, 600));
                    const res = await checkNptu(valEisbn, 'k');
                    if (res.found) { foundMatch = 'eISBN'; collectionQuantity = res.quantity; nptuUrl = res.url; }
                }

                // 1.3 Try searching by pISBN
                if (!foundMatch && valPisbn && valPisbn.length >= 8) {
                    processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): [查pISBN] ${valPisbn}`;
                    await new Promise(r => setTimeout(r, 600));
                    const res = await checkNptu(valPisbn, 'k');
                    if (res.found) { foundMatch = 'pISBN'; collectionQuantity = res.quantity; nptuUrl = res.url; }
                }
                
                // 1.5 Try searching by EISSN
                if (!foundMatch && valEissn && valEissn.length >= 8) {
                    processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): [查EISSN] ${valEissn}`;
                    await new Promise(r => setTimeout(r, 600));
                    const res = await checkNptu(valEissn, 'k');
                    if (res.found) { foundMatch = 'EISSN'; collectionQuantity = res.quantity; nptuUrl = res.url; }
                }
                
                // Remove all punctuation/special characters for title search
                // \p{L} is any letter (including Chinese), \p{N} is any number
                let titleNoSpecial = title ? title.replace(/[^\p{L}\p{N}]/gu, ' ').replace(/\s+/g, ' ').trim() : '';

                // 2. Try Title (with special chars replaced by space)
                if (!foundMatch && titleNoSpecial) {
                    processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): [查書名] ${titleNoSpecial}`;
                    const res = await checkNptu(titleNoSpecial, 'k');
                    if (res.found) { foundMatch = '書名'; collectionQuantity = res.quantity; nptuUrl = res.url; }
                }
                
                // 3. Try "Cleaned" Title (remove subtitles etc, taking only part before typical subtitle separators)
                if (!foundMatch && title) {
                    // Split the ORIGINAL title by colon/parens to get the main title part
                    let corePart = title.split(/：|:|（|\(| - |\/|／|《|〈|\[|【/)[0].trim();
                    // Clean special chars out of that core part too
                    let cleanTitle = corePart.replace(/[^\p{L}\p{N}]/gu, ' ').replace(/\s+/g, ' ').trim();
                    
                    if (cleanTitle && cleanTitle.length > 1 && cleanTitle !== titleNoSpecial) {
                        processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): [模糊書名] ${cleanTitle}`;
                        // Delay again to protect from rate limit since we make another request
                        await new Promise(r => setTimeout(r, 600));
                        const res = await checkNptu(cleanTitle, 'k');
                        if (res.found) { foundMatch = '書名(模糊)'; collectionQuantity = res.quantity; nptuUrl = res.url; }
                    }
                }
                
                if (foundMatch) {
                    isDuplicate = `是 (${foundMatch})`;
                    rowStatus = "duplicate";
                    statusText = isDuplicate;
                } else {
                    isDuplicate = "否"; // Only "否" if all 3 attempts failed
                    rowStatus = "not-found";
                    statusText = "無複本";
                    collectionQuantity = "0";
                }
            } catch (err) {
                console.error("查核失敗 " + (title || valIsbn), err);
                rowStatus = "error";
                let shortErr = err.message ? err.message.substring(0, 50) : "查詢異常";
                statusText = "錯誤: " + shortErr;
                isDuplicate = "查詢失敗";
                collectionQuantity = "錯誤";
            }

            processStatus.textContent = `處理中 (${i + 1}/${excelData.length}): [查寫真/限制級] ${displayBarcode}`;
            let isbnsToCheck = [];
            if (valIsbn) isbnsToCheck.push(valIsbn);
            if (valEisbn) isbnsToCheck.push(valEisbn);
            if (valPisbn) isbnsToCheck.push(valPisbn);

            try {
                if (isbnsToCheck.length === 0) {
                    specialProperty = await checkSpecialProperty("", title);
                } else {
                    let lastResult = "無";
                    for (let idx = 0; idx < isbnsToCheck.length; idx++) {
                        const currentIsbn = isbnsToCheck[idx];
                        try {
                            if (idx > 0 || i > 0) await new Promise(r => setTimeout(r, 600));
                            let res = await checkSpecialProperty(currentIsbn, title);
                            if (res !== "無" && res !== "查詢失敗") {
                                lastResult = res;
                                break; // Found photo book or restricted!
                            }
                            if (res === "無") {
                                lastResult = "無";
                            }
                        } catch (e) {
                            console.error(`查詢 ISBN ${currentIsbn} 特殊屬性失敗:`, e);
                            if (lastResult === "無") lastResult = "查詢失敗";
                        }
                    }
                    specialProperty = lastResult;
                }
            } catch (err) {
                specialProperty = "查詢失敗";
            }

            // Append field to data
            row["寫真書或限制級圖書"] = specialProperty;
            row["是否為複本"] = isDuplicate;
            row["館藏數量"] = collectionQuantity;
            if (nptuUrl) row["系統連結"] = nptuUrl;
            processedData.push(row);
            createOrUpdateRow(i, displayBarcode, title, author, publisher, pubYear, specialProperty, rowStatus, statusText, collectionQuantity, nptuUrl);
        }

        isProcessing = false;
        stopBtn.style.display = 'none';
        if (processStatus.textContent.indexOf('手動停止') === -1) {
            processStatus.textContent = `處理完成！所有明細已確認完畢。`;
            progressBar.style.width = '100%';
        }
        downloadBtn.disabled = false;
        resetBtn.style.display = 'inline-flex';
    }

    function createOrUpdateRow(index, isbn, title, author, publisher, pubYear, specialProperty, statusClass, statusText, quantity = "-", link = "") {
        let tr = document.getElementById(`row-${index}`);
        if (!tr) {
            tr = document.createElement('tr');
            tr.id = `row-${index}`;
            resultsTbody.appendChild(tr);
            
            const wrapper = document.querySelector('.table-wrapper');
            wrapper.scrollTop = wrapper.scrollHeight;
        }

        let icon = '';
        if (statusClass === 'fetching') icon = '<i class="ph ph-spinner ph-spin"></i>';
        else if (statusClass === 'duplicate') icon = '<i class="ph-fill ph-check-circle"></i>';
        else if (statusClass === 'not-found') icon = '<i class="ph-fill ph-minus-circle"></i>';
        else if (statusClass === 'error') icon = '<i class="ph-fill ph-warning-circle"></i>';

        const displayTitle = title.length > 25 ? title.substring(0, 25) + '...' : title;
        const displayAuthor = author.length > 15 ? author.substring(0, 15) + '...' : author;
        const displayPublisher = publisher.length > 15 ? publisher.substring(0, 15) + '...' : publisher;
        const displayPubYear = pubYear ? pubYear : '-';

        // UI 標紅邏輯
        const isDuplicateRed = statusText.includes('是');
        const isQtyRed = !isNaN(parseInt(quantity)) && parseInt(quantity) > 0;

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td style="font-family:monospace">${isbn || '-'}</td>
            <td>${displayTitle || '-'}</td>
            <td>${displayAuthor || '-'}</td>
            <td>${displayPublisher || '-'}</td>
            <td>${displayPubYear}</td>
            <td><span class="badge" style="background:#f0f0f0;padding:2px 6px;border-radius:4px;font-size:0.9em;color:${(specialProperty.includes('限制級') || specialProperty.includes('寫真書'))?'#d32f2f':'#666'}">${specialProperty || '-'}</span></td>
            <td><span class="status status-${statusClass}" style="${isDuplicateRed ? 'color: #d32f2f; background: rgba(211,47,47,0.1);' : ''}">${icon} ${statusText}</span></td>
            <td style="text-align:center;font-weight:bold;color:${isQtyRed ? '#d32f2f' : '#4f46e5'};">${quantity}</td>
            <td style="text-align:center;">${link ? `<a href="${link}" target="_blank" style="color:#4f46e5;text-decoration:none;"><i class="ph-bold ph-link"></i></a>` : '-'}</td>
        `;
    }

    // Helper for timeout fetch
    async function fetchWithTimeout(url, timeoutMs) {
        const controller = new AbortController();
        const id = setTimeout(() => controller.abort(), timeoutMs);
        try {
            const res = await fetch(url, { signal: controller.signal });
            clearTimeout(id);
            return res;
        } catch (e) {
            clearTimeout(id);
            throw e;
        }
    }

    async function checkNptu(query, typeStr) {
        // Build WEBPAC search URL (must include &si=1&view=d and use t0=k to correctly parse the page)
        const searchUrl = `https://webpac.nptu.edu.tw/webpac/search.cfm?m=ss&k0=${encodeURIComponent(query)}&t0=${typeStr}&c0=and&s0=0&w=0&list_num=10&current_page=1&si=1&view=d`;
        
        // Strategy array for bypassing CORS
        const proxies = [
            searchUrl, // Try direct first! (if user uses CORS unblocker)
            `https://api.allorigins.win/raw?url=${encodeURIComponent(searchUrl)}`,
            `https://corsproxy.io/?${encodeURIComponent(searchUrl)}`
        ];

        let lastErr = "";
        for (const proxy of proxies) {
            try {
                const res = await fetchWithTimeout(proxy, 10000);
                if (!res.ok) {
                    lastErr = `${res.status} HTTP Error`;
                    continue;
                }
                
                const redirectedUrl = decodeURIComponent(res.url || "");
                const html = await res.text();
                
                let quantityStr = null;
                
                // 為了避免 HTML 標籤或是 &nbsp; 空格實體影響比對，先清理字串
                let cleanText = html.replace(/<[^>]+>/g, '').replace(/&nbsp;/ig, ' ');
                
                // 涵蓋 "館藏(3)", "館藏量 (2)", "館藏數量(1)" 等各種括號表達式
                let qtyMatch = cleanText.match(/館藏(?:量|數量)?\s*\(\s*(\d+)\s*\)/);
                
                if (qtyMatch) {
                    quantityStr = qtyMatch[1];
                } else {
                    // 備案：涵蓋 "館藏數: 1", "館藏: 1", "館藏量 2" 等狀況
                    let altMatch = cleanText.match(/館藏(?:量|數|數量)?\s*[:：]?\s*(\d+)/);
                    if (altMatch) quantityStr = altMatch[1];
                }

                // If corsproxy follows a 302 redirect natively to the details page, or html indicates results
                if (
                    redirectedUrl.includes("content.cfm?mid=") || 
                    redirectedUrl.includes("marc.cfm?mid=") ||
                    html.includes("content.cfm?mid=") || 
                    html.includes("Search Results") ||
                    html.includes("圖書資訊") ||
                    html.includes("書目資料")
                ) {
                    // Confirm it actually found records, not just "0 Search Results"
                    if (html.includes("沒有找到任何符合的資料") || html.includes("查無資料") || html.includes("0 筆結果")) {
                        return { found: false, quantity: "0", url: "" };
                    }
                    return { found: true, quantity: quantityStr || "1+", url: searchUrl };
                }
                return { found: false, quantity: "0", url: "" }; 
            } catch (err) {
                // Ignore and try next proxy
                lastErr = err.message;
            }
        }
        throw new Error("連線被拒，請安裝並開啟「Allow CORS」擴充功能再重試 (原錯誤: " + (lastErr || "代理伺服器無回應") + ")");
    }

    async function checkSpecialProperty(isbn, title) {
        let result = [];
        if (title && (title.includes("寫真") || title.includes("限制級") || title.includes("18禁"))) {
            if (title.includes("寫真")) result.push("寫真書");
            if (title.includes("限制級") || title.includes("18禁")) result.push("限制級");
        }

        if (!isbn || isbn.length < 8) {
            return result.length > 0 ? result.join("/") : '無'; // 若無有效 ISBN/ISSN 也當作查無此書處理
        }

        let querySuccess = false;
        
        // 依要求使用三民書局
        const urls = [
            `https://www.sanmin.com.tw/search?isbn=${encodeURIComponent(isbn)}`
        ];

        for (const url of urls) {
            const proxies = [
                url, // 加入直接連線 (若使用者有裝 CORS 解鎖套件會優先成功且最快)
                `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
                `https://corsproxy.io/?${encodeURIComponent(url)}`
            ];

            for (const proxy of proxies) {
                try {
                    const res = await fetchWithTimeout(proxy, 10000); // 延長至 10 秒避免代理商回應太慢
                    if (!res.ok) continue;

                    const rawHtml = await res.text();
                    
                    // 確保真的抓到了網站，而不是代理伺服器的錯誤頁面
                    if (!rawHtml.includes("三民") && !rawHtml.includes("sanmin") && !rawHtml.includes("商品") && !rawHtml.includes("查詢")) {
                        continue;
                    }

                    querySuccess = true;
                    const htmlStr = rawHtml.replace(/\s+/g, '');
                    
                    // 檢查是否查無此書
                    if (htmlStr.includes("找不到符合的商品") || 
                        htmlStr.includes("查無相關商品") || 
                        htmlStr.includes("沒有找到任何符合的資料") || 
                        htmlStr.includes("沒有符合的搜尋結果") || 
                        htmlStr.includes("您查詢的關鍵字沒有符合的商品") ||
                        htmlStr.includes("抱歉，沒有找到")) {
                        
                        // 若查無此書，且標題本身也沒寫真/限制級關鍵字，則填"無"
                        if (result.length === 0) return "無";
                        else return result.join("/");
                    }

                    // 從三民書局的畫面判斷：圖片上有「限制級」、「18+」、「限制級商品」的標籤
                    if (!result.includes("限制級")) {
                        if (
                            htmlStr.includes("限制級商品") || 
                            htmlStr.includes(">18+<") ||
                            htmlStr.includes("18禁") ||
                            htmlStr.includes("級別:限制級") || 
                            htmlStr.includes("級別：限制級") ||
                            />[^<]*(限制級|十八歲以下禁止|未滿18歲)[^>]*</i.test(rawHtml)
                        ) {
                            result.push("限制級");
                        }
                    }
                    
                    if (!result.includes("寫真書")) {
                        if (
                            htmlStr.includes(">寫真<") || 
                            htmlStr.includes("分類:寫真") || 
                            htmlStr.includes("分類：寫真") ||
                            htmlStr.includes("寫真集") ||
                            htmlStr.includes("寫真書") ||
                            htmlStr.includes("偶像寫真") ||
                            />[^<]*(寫真集|寫真書|偶像寫真)[^>]*</i.test(rawHtml)
                        ) {
                            result.push("寫真書");
                        }
                    }

                    break; // 成功抓取該 URL 的資料就可脫離 proxy 迴圈
                } catch(e) { }
            }
            if (querySuccess) break;
        }
        
        return result.length > 0 ? result.join("/") : (querySuccess ? '否' : '查詢失敗');
    }

    function exportExcel() {
        if (processedData.length === 0) return;
        
        const newWorkbook = XLSX.utils.book_new();

        // 將資料依照 __sheetName 分組
        const sheetsData = {};
        for (const row of processedData) {
            const sn = row['__sheetName'] || 'Sheet1';
            if (!sheetsData[sn]) sheetsData[sn] = [];
            const exportRow = { ...row };
            delete exportRow['__sheetName']; // 移出內部分頁標記
            delete exportRow['__SYS_ISBN'];
            delete exportRow['__SYS_EISSN'];
            delete exportRow['__SYS_TITLE'];
            delete exportRow['__SYS_AUTHOR'];
            delete exportRow['__SYS_PUBLISHER'];
            delete exportRow['__SYS_PUBYEAR'];
            delete exportRow['__SYS_SEQ_KEY'];
            sheetsData[sn].push(exportRow);
        }

        // 為每個分組建立 Sheet
        for (const sheetName in sheetsData) {
            const data = sheetsData[sheetName];
            const newWorksheet = XLSX.utils.json_to_sheet(data);
            
            if (newWorksheet['!ref']) {
                const range = XLSX.utils.decode_range(newWorksheet['!ref']);
                let headers = {};
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ c: C, r: range.s.r });
                    if (newWorksheet[cellRef]) {
                        headers[C] = newWorksheet[cellRef].v;
                    }
                }

                for (let R = range.s.r + 1; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
                        const cell = newWorksheet[cellRef];
                        if (!cell || !cell.v) continue;

                        const header = headers[C];
                        const val = String(cell.v);
                        let shouldBeRed = false;

                        if (header === '是否為複本' && val.includes('是')) {
                            shouldBeRed = true;
                        }
                        if (header === '寫真書或限制級圖書' && (val.includes('限制級') || val.includes('寫真書'))) {
                            shouldBeRed = true;
                        }
                        if (header === '館藏數量' && !isNaN(parseInt(val)) && parseInt(val) > 0) {
                            shouldBeRed = true;
                        }

                        if (shouldBeRed) {
                            cell.s = {
                                font: { color: { rgb: "FF0000" }, bold: true }
                            };
                        }
                        
                        // 設定「系統連結」欄位為可點擊超連結的樣式
                        if (header === '系統連結' && val && val.startsWith('http')) {
                            cell.l = { Target: val };
                            cell.s = cell.s || {};
                            cell.s.font = cell.s.font || {};
                            cell.s.font.color = { rgb: "0563C1" };
                            cell.s.font.underline = true;
                        }
                    }
                }
            }
            
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
        }
        
        XLSX.writeFile(newWorkbook, "Checked_Duplicate_Books.xlsx");
    }
});
