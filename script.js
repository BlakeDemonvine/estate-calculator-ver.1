// --- 1. 資料與全域設定 ---
Chart.register(ChartDataLabels);

// 儲存多個圖表實體
const charts = {}; 
// 顏色盤
const colorPalette = ['#FF9AA2', '#E27D9A', '#B0558F', '#7D3C98', '#4A2399', '#1D1E8F'];
// 換算係數
const PING_FACTOR = 0.3025;
// Modal 當前編輯 ID
let currentEditingId = null;

let content = document.getElementById('canvas-area');

// --- 2. 初始化 ---
document.addEventListener("DOMContentLoaded", function() {
    const targetIds = ['A', 'B', 'C', 'D', 'E', 'F'];
    const items = targetIds
        .map(id => document.getElementById(id))
        .filter(item => item !== null);
        
    items.forEach(item => {
        item.addEventListener('click', function() {
            items.forEach(el => el.classList.remove('active'));
            this.classList.add('active');
            show(this.id);
        });
    });

    show('A'); // 預設顯示 A
});

// --- 3. 頁面切換邏輯 ---
function show(input){
    content.innerHTML = '';
    content.className = 'canvas-area'; // 重置 class
    document.getElementById('globalSummary').classList.add('hidden-bar');
    
    // 銷毀舊圖表
    Object.keys(charts).forEach(id => {
        if (charts[id]) {
            charts[id].destroy();
            delete charts[id];
        }
    });

    const landSegment = basics['地段'] || '未命名專案';
    document.getElementById('title').value = landSegment;
    document.title = landSegment + " - 磚家計算";

    // 在 script.js 中修改 show 函數的 A 部分
    if(input === 'A'){
        content.style.display = 'flex';
        content.style.flexDirection = 'column';
        content.style.alignItems = 'center';
        content.style.justifyContent = 'center';

        let fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.multiple = true;
        fileInput.accept = 'application/pdf';
        fileInput.style.display = 'none';

        let cloud = document.createElement('i');
        cloud.className = 'fa-solid fa-cloud-arrow-up';
        cloud.style.cssText = 'color:var(--accent-purple); font-size:150px; margin-bottom:20px; cursor:pointer;';
        
        let btn = document.createElement('button');
        btn.className = 'btn-export';
        btn.textContent = '匯入多份謄本';

        // 觸發上傳
        cloud.onclick = () => fileInput.click();
        btn.onclick = () => fileInput.click();

        fileInput.onchange = async (e) => {
            const files = e.target.files;
            if (files.length === 0) return;

            btn.disabled = true;
            btn.textContent = "正在解析...";

            for (let file of files) {
                // 這裡呼叫剛剛定義在外部的 function
                await processTranscriptPDF(file);
            }

            alert("解析完成！");
            btn.disabled = false;
            btn.textContent = '匯入多份謄本';
            show('B'); // 切換到清冊頁面
        };

        content.appendChild(fileInput);
        content.appendChild(cloud);
        content.appendChild(btn);
    }
    else if(input === 'B'){
        // --- B 頁面: 清冊 (加入搜尋功能) ---
        content.style.display = ''; 

        // 1. 建立搜尋列
        let searchContainer = document.createElement('div');
        searchContainer.style.cssText = 'padding: 20px 40px 0 40px; display: flex; justify-content: center;';
        searchContainer.innerHTML = `
            <div style="position: relative; width: 100%; max-width: 400px;">
                <i class="fa-solid fa-magnifying-glass" style="position: absolute; left: 15px; top: 50%; transform: translateY(-50%); color: #888;"></i>
                <input type="text" id="searchB" placeholder="搜尋地號..." 
                    style="width: 100%; padding: 12px 12px 12px 40px; border-radius: 50px; border: 1px solid #ccc; font-size: 16px; outline: none;">
            </div>
        `;
        content.appendChild(searchContainer);
        
        // 2. 建立 Grid 容器
        let gridContainer = document.createElement('div');
        gridContainer.classList.add('chart-grid-container');
        gridContainer.id = 'gridB';
        content.appendChild(gridContainer);

        // 3. 搜尋邏輯
        document.getElementById('searchB').addEventListener('input', function(e) {
            const val = e.target.value.toLowerCase();
            const cards = gridContainer.querySelectorAll('.chart-card');
            cards.forEach(card => {
                const title = card.querySelector('.chart-title').innerText.toLowerCase();
                if(title.includes(val)) {
                    card.style.display = 'flex';
                } else {
                    card.style.display = 'none';
                }
            });
        });

        for(let num in data){
            //if (num === '地段') continue;

            // 產生 HTML
            let cardHTML = createChartBlock(`chart_${num}`, `地號：${num}`);
            gridContainer.insertAdjacentHTML('beforeend', cardHTML);
            
            // 初始化圖表
            initChart(`chart_${num}`);
            
            // 加入資料
            const owners = data[num]['所有權人'];
            if(owners && Array.isArray(owners)){
                for(let i = 0 ; i < owners.length ; i++){
                    let areaVal = data[num]['土地面積']['面積'];
                    let scopes = data[num]['土地面積']['權利範圍'];
                    // 處理單一數值或陣列
                    let scopeVal = Array.isArray(scopes) ? scopes[i] : scopes;
                    
                    addPerson(`chart_${num}`, owners[i], areaVal * scopeVal);
                }
            }
        }
        updateGlobalSummary();
        document.getElementById('globalSummary').classList.remove('hidden-bar');
    }
    else if(input === 'C'){
        // --- C 頁面: 增值稅歸戶表 (以「人」為單位) ---
        content.style.display = '';
        
        // 1. 搜尋列
        let searchContainer = document.createElement('div');
        searchContainer.style.cssText = 'padding: 20px 40px 0 40px; display: flex; justify-content: center;';
        searchContainer.innerHTML = `
             <div style="position: relative; width: 100%; max-width: 400px;">
                <i class="fa-solid fa-user" style="position: absolute; left: 15px; top: 50%; transform: translateY(-50%); color: #888;"></i>
                <input type="text" id="searchC" placeholder="搜尋所有權人..." 
                    style="width: 100%; padding: 12px 12px 12px 40px; border-radius: 50px; border: 1px solid #ccc; font-size: 16px; outline: none;">
            </div>
        `;
        content.appendChild(searchContainer);

        let gridContainer = document.createElement('div');
        gridContainer.classList.add('chart-grid-container');
        // 調整 Grid 為單欄或雙欄以適應詳細表格
        gridContainer.style.gridTemplateColumns = 'repeat(auto-fill, minmax(500px, 1fr))';
        content.appendChild(gridContainer);
        document.getElementById('globalSummary').classList.add('hidden-bar');

        // 2. 資料重組 (Data Aggregation): Land-based -> Person-based
        let personData = {}; // { '王小明': { lands: [], totalPing: 0, totalTaxSelf: 0, ... } }

        for (let landId in data) {
            //if (landId === '地段') continue;
            let record = data[landId];
            let owners = record['所有權人'] || [];
            
            for (let i = 0; i < owners.length; i++) {
                let name = owners[i];
                if (!personData[name]) {
                    personData[name] = {
                        name: name,
                        lands: [],
                        totalAreaM2: 0,
                        totalPing: 0,
                        totalTaxSelf: 0,
                        totalTaxGen: 0
                    };
                }

                // 取值
                let area = record['土地面積']['面積'] || 0;
                let scopes = record['土地面積']['權利範圍'];
                let scope = Array.isArray(scopes) ? scopes[i] : scopes;
                let dates = record['年月'] || [];
                let prevVals = record['公告現值'] || [];
                let taxSelfs = record['增值稅預估(自用)'] || [];
                let taxGens = record['增值稅預估(一般)'] || [];
                let currVal102 = record['當期公告現值'] || 0;

                // 計算
                let heldM2 = area * scope;
                let heldPing = heldM2 * PING_FACTOR;
                let taxSelf = parseFloat(Array.isArray(taxSelfs) ? taxSelfs[i] : taxSelfs) || 0;
                let taxGen = parseFloat(Array.isArray(taxGens) ? taxGens[i] : taxGens) || 0;

                // 存入個人明細
                personData[name].lands.push({
                    id: landId,
                    scope: scope,
                    heldPing: heldPing,
                    date: Array.isArray(dates) ? dates[i] : dates,
                    prevVal: Array.isArray(prevVals) ? prevVals[i] : prevVals,
                    currVal: currVal102,
                    taxSelf: taxSelf,
                    taxGen: taxGen
                });

                // 累加總計
                personData[name].totalAreaM2 += heldM2;
                personData[name].totalPing += heldPing;
                personData[name].totalTaxSelf += taxSelf;
                personData[name].totalTaxGen += taxGen;
            }
        }

        // 3. 渲染卡片
        Object.values(personData).forEach(person => {
            let card = document.createElement('div');
            card.className = 'chart-card';
            card.style.alignItems = 'stretch'; // 讓內容滿寬
            card.style.height = 'auto'; // 高度自動

            // 標題
            let header = `
                <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:2px solid #f0f0f0; padding-bottom:15px; margin-bottom:15px;">
                    <h3 style="font-size:22px; font-weight:700; color:#333; margin:0;">${person.name}</h3>
                    <span style="background:var(--accent-purple); color:white; padding:4px 12px; border-radius:15px; font-size:12px;">持有 ${person.lands.length} 筆</span>
                </div>
            `;

            // 表格內容
            let rowsHTML = person.lands.map(land => `
                <tr style="border-bottom:1px solid #eee;">
                    <td style="padding:10px; font-weight:bold; color:#0056b3;">${land.id}</td>
                    <td style="padding:10px; text-align:center;">${(land.scope).toFixed(4)}</td> <td style="padding:10px; text-align:right;">${land.heldPing.toFixed(2)}</td>
                    <td style="padding:10px; text-align:right; color:#888;">${parseFloat(land.currVal).toLocaleString()}</td>
                    <td style="padding:10px; text-align:right; color:#d93025;">${Math.round(land.taxSelf).toLocaleString()}</td>
                    <td style="padding:10px; text-align:right;">${Math.round(land.taxGen).toLocaleString()}</td>
                </tr>
            `).join('');

            // 表格結構
            let table = `
                <div style="overflow-x:auto;">
                    <table style="width:100%; border-collapse:collapse; font-size:14px;">
                        <thead>
                            <tr style="background:#f9f9f9; color:#666;">
                                <th style="padding:8px; text-align:left;">地號</th>
                                <th style="padding:8px; text-align:center;">權利範圍</th>
                                <th style="padding:8px; text-align:right;">持分(坪)</th>
                                <th style="padding:8px; text-align:right;">當期現值</th>
                                <th style="padding:8px; text-align:right;">增值稅(自)</th>
                                <th style="padding:8px; text-align:right;">增值稅(般)</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rowsHTML}
                        </tbody>
                    </table>
                </div>
            `;

            // 總計區塊
            let footer = `
                <div style="margin-top:20px; background:#f0f7ff; padding:15px; border-radius:8px; display:flex; justify-content:space-around; align-items:center;">
                    <div style="text-align:center;">
                        <div style="font-size:12px; color:#666;">總持分坪數</div>
                        <div style="font-size:18px; font-weight:bold; color:#333;">${person.totalPing.toFixed(2)}</div>
                    </div>
                    <div style="width:1px; height:30px; background:#ddd;"></div>
                    <div style="text-align:center;">
                        <div style="font-size:12px; color:#666;">增值稅合計(自)</div>
                        <div style="font-size:18px; font-weight:bold; color:#d93025;">$${Math.round(person.totalTaxSelf).toLocaleString()}</div>
                    </div>
                    <div style="width:1px; height:30px; background:#ddd;"></div>
                     <div style="text-align:center;">
                        <div style="font-size:12px; color:#666;">增值稅合計(般)</div>
                        <div style="font-size:18px; font-weight:bold; color:#333;">$${Math.round(person.totalTaxGen).toLocaleString()}</div>
                    </div>
                </div>
            `;

            card.innerHTML = header + table + footer;
            gridContainer.appendChild(card);
        });

        // 搜尋監聽
        document.getElementById('searchC').addEventListener('input', function(e) {
            const val = e.target.value.toLowerCase();
            const cards = gridContainer.querySelectorAll('.chart-card');
            cards.forEach(card => {
                const name = card.querySelector('h3').innerText.toLowerCase();
                if(name.includes(val)) {
                    card.style.display = 'flex';
                } else {
                    card.style.display = 'none';
                }
            });
        });
    }
    else if(input === 'D'){
        // --- D 頁面: 土地歸戶表 + 參數輸入 ---
        content.style.display = '';
        document.getElementById('globalSummary').classList.add('hidden-bar');

        // 1. 讀取 basics 參數
        // basics['土地歸戶表']['預估產權面積(坪)'] 是 [factor1, factor2, factor3]
        // basics['土地歸戶表']['合建分取'] 是 ratio
        let params = basics['土地歸戶表']['預估產權面積(坪)'] || [0, 0, 0];
        let jointRatio = basics['土地歸戶表']['合建分取'] || 0;

        // 2. 建立上方控制面板 (Inputs)
        let controlPanel = document.createElement('div');
        controlPanel.style.cssText = 'background: white; padding: 20px; border-radius: 12px; margin: 20px 40px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); display: flex; flex-wrap: wrap; gap: 20px; align-items: center; justify-content: center;';
        
        controlPanel.innerHTML = `
            <div style="display:flex; align-items:center; gap:10px;">
                <strong style="color:var(--accent-purple);">預估產權參數：</strong>
                <input type="number" id="param1" value="${params[0]}" step="0.01" style="width:80px; padding:8px; border:1px solid #ddd; border-radius:4px;">
                <span>x</span>
                <input type="number" id="param2" value="${params[1]}" step="0.01" style="width:80px; padding:8px; border:1px solid #ddd; border-radius:4px;">
                <span>x</span>
                <input type="number" id="param3" value="${params[2]}" step="0.01" style="width:80px; padding:8px; border:1px solid #ddd; border-radius:4px;">
            </div>
            <div style="width:1px; height:30px; background:#eee;"></div>
            <div style="display:flex; align-items:center; gap:10px;">
                <strong style="color:var(--accent-purple);">合建分取比率：</strong>
                <input type="number" id="jointParam" value="${jointRatio}" step="0.01" style="width:80px; padding:8px; border:1px solid #ddd; border-radius:4px;">
            </div>
        `;
        content.appendChild(controlPanel);

        // 3. 建立 Grid
        let gridContainer = document.createElement('div');
        gridContainer.classList.add('chart-grid-container');
        gridContainer.id = 'gridD';
        content.appendChild(gridContainer);

        // 定義渲染函數 (因為輸入改變時要重繪)
        function renderCards() {
            gridContainer.innerHTML = '';
            
            // 取得最新參數
            const p1 = parseFloat(document.getElementById('param1').value) || 0;
            const p2 = parseFloat(document.getElementById('param2').value) || 0;
            const p3 = parseFloat(document.getElementById('param3').value) || 0;
            const jRatio = parseFloat(document.getElementById('jointParam').value) || 0;

            // 更新回全域變數 basics
            basics['土地歸戶表']['預估產權面積(坪)'] = [p1, p2, p3];
            basics['土地歸戶表']['合建分取'] = jRatio;

            if(typeof window.saveProjectToLocal === 'function') window.saveProjectToLocal();

            // 資料聚合 (Person-based)
            let personData = {};
            for (let landId in data) {
                //if (landId === '地段') continue;
                let record = data[landId];
                let owners = record['所有權人'] || [];
                
                for (let i = 0; i < owners.length; i++) {
                    let name = owners[i];
                    if (!personData[name]) {
                        personData[name] = {
                            name: name,
                            totalHeldPing: 0,
                            lands: [] // 用來計算細節(如果需要顯示的話)
                        };
                    }
                    let area = record['土地面積']['面積'] || 0;
                    let scopes = record['土地面積']['權利範圍'];
                    let scope = Array.isArray(scopes) ? scopes[i] : scopes;
                    
                    let heldPing = area * scope * PING_FACTOR;
                    personData[name].totalHeldPing += heldPing;
                }
            }

            // 產生卡片
            Object.values(personData).forEach(person => {
                // 依照 excel.js SheetC 邏輯:
                // G (總持分坪數) = totalHeldPing
                // H (預估產權面積) = G * p1 * p2 * p3
                // K (合建分取) = H * jRatio

                let estPropArea = person.totalHeldPing * p1 * p2 * p3;
                let jointAlloc = estPropArea * jRatio;

                let card = document.createElement('div');
                card.className = 'chart-card';
                card.style.height = 'auto'; 
                card.style.background = 'linear-gradient(to bottom right, #ffffff, #fdfbff)';
                
                card.innerHTML = `
                    <h3 style="width:100%; border-bottom:1px solid #eee; padding-bottom:10px; margin-bottom:15px; color:#333;">${person.name}</h3>
                    
                    <div style="width:100%; display:flex; flex-direction:column; gap:12px;">
                        <div class="info-row">
                            <span class="info-label">總持分面積 (坪)</span>
                            <span class="info-value" style="font-size:18px;">${person.totalHeldPing.toFixed(2)}</span>
                        </div>
                        
                        <div style="background:#e8f5e9; padding:10px; border-radius:6px; border:1px solid #c8e6c9;">
                            <div class="info-label" style="color:#2e7d32; margin-bottom:4px;">預估產權面積 (坪)</div>
                            <div class="info-value" style="color:#1b5e20; font-size:20px;">${estPropArea.toFixed(2)}</div>
                            <div style="font-size:10px; color:#666; margin-top:4px;">公式: 總持分 x ${p1} x ${p2} x ${p3}</div>
                        </div>

                        <div style="background:#fff3e0; padding:10px; border-radius:6px; border:1px solid #ffe0b2;">
                            <div class="info-label" style="color:#ef6c00; margin-bottom:4px;">合建分取 (坪)</div>
                            <div class="info-value" style="color:#e65100; font-size:20px;">${jointAlloc.toFixed(2)}</div>
                             <div style="font-size:10px; color:#666; margin-top:4px;">公式: 產權面積 x ${jRatio}</div>
                        </div>
                    </div>
                `;
                gridContainer.appendChild(card);
            });
        }

        // 綁定事件：輸入框變動時即時計算
        ['param1', 'param2', 'param3', 'jointParam'].forEach(id => {
            document.getElementById(id).addEventListener('input', renderCards);
        });

        // 初始渲染
        renderCards();
    }
}

function createChartBlock(id, title) {
    const rawId = id.replace('chart_', '');
    return `
        <div class="chart-card" data-id="${rawId}">
            <h3 class="chart-title">${title}</h3>
            <div class="chart-wrapper" style="position: relative; width: 100%; aspect-ratio: 1/1; max-width: 300px;"> 
                <canvas id="canvas_${id}"></canvas>
                <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); text-align: center; pointer-events: none;">
                    <div style="font-size: 13px; color: #888; margin-bottom: 4px;">持分總計</div>
                    <div id="val_${id}" style="font-size: 24px; font-weight: 800; color: #2c3e50;">0.0</div>
                </div>
            </div>
            <div style="margin-top: auto; padding-top: 20px; width: 100%; display: flex; justify-content: center; gap: 10px;">
                <button class="btn-edit" onclick="openEditModal('${rawId}')">
                    <i class="fa-solid fa-pen-to-square"></i> 編輯
                </button>
                <button class="btn-edit" style="color: #d93025; border-color: #d93025;" onclick="deleteLand('${rawId}')">
                    <i class="fa-solid fa-trash"></i> 刪除
                </button>
            </div>
        </div>
    `;
}

// 請將這個新函式放在 script.js 中的任意空位 (例如檔案最下方)
function deleteLand(id) {
    if (confirm(`確定要刪除地號「${id}」嗎？此操作將無法復原。`)) {
        // 從資料結構中刪除
        delete data[id];
        
        // 從圖表庫中移除
        if (charts[`chart_${id}`]) {
            charts[`chart_${id}`].destroy();
            delete charts[`chart_${id}`];
        }
        
        // 儲存進 Local Storage
        if (typeof window.saveProjectToLocal === 'function') {
            window.saveProjectToLocal();
        }
        
        // 重新整理目前的分頁畫面
        const activeTab = document.querySelector('.nav-item.active');
        if (activeTab) {
            show(activeTab.id);
        }
    }
}

// --- 4. Chart.js 邏輯 (保持不變) ---
function initChart(id) {
    const ctx = document.getElementById(`canvas_${id}`);
    if (!ctx) return;

    const config = {
        type: 'doughnut',
        data: {
            labels: [], 
            datasets: [{
                data: [], 
                backgroundColor: [], 
                borderWidth: 2,
                borderColor: '#ffffff', 
                hoverOffset: 10 
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '55%', 
            layout: { padding: 20 },
            plugins: {
                legend: { display: false },
                tooltip: { enabled: true },
                datalabels: {
                    color: '#ffffff', 
                    font: { weight: 'bold', size: 14 },
                    formatter: (value, context) => {
                        const name = context.chart.data.labels[context.dataIndex];
                        const datapoints = context.chart.data.datasets[0].data;
                        const total = datapoints.reduce((a, b) => a + b, 0);
                        const percentage = ((value / total) * 100).toFixed(1) + '%';
                        return name + '\n' + percentage;
                    },
                    textAlign: 'center', anchor: 'center', align: 'center'
                }
            },
            animation: { animateScale: true, animateRotate: true }
        }
    };
    charts[id] = new Chart(ctx, config);
}

function addPerson(chartId, name, amount) {
    const targetChart = charts[chartId];
    if (!targetChart) return;

    targetChart.data.labels.push(name);
    targetChart.data.datasets[0].data.push(amount);
    const colorIndex = (targetChart.data.labels.length - 1) % colorPalette.length;
    targetChart.data.datasets[0].backgroundColor.push(colorPalette[colorIndex]);
    targetChart.update();
    updateCenterTotal(chartId);
}

function updateCenterTotal(chartId) {
    const targetChart = charts[chartId];
    if (!targetChart) return;
    const data = targetChart.data.datasets[0].data;
    const total = data.reduce((a, b) => a + b, 0);
    const centerEl = document.getElementById(`val_${chartId}`);
    if (centerEl) centerEl.innerText = total.toFixed(1);
}

// --- 6. 新增地號與編輯 Modal 功能 (保持不變，略作整理) ---

function newNum() {
    if (document.getElementById('newNumModal')) return;

    const modalHTML = `
        <div id="newNumModal" class="modal-overlay">
            <div class="modal-content" style="width: 380px; height: auto; overflow: visible;">
                <div class="modal-header">
                    <h3>新增地號</h3>
                    <button class="close-btn" onclick="closeNewNumModal()">&times;</button>
                </div>
                <div class="modal-body" style="padding: 25px 24px; overflow: visible;">
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label class="field-label" style="font-size: 14px; margin-bottom: 8px;"><span style="color:#d93025;">*</span> 地號</label>
                        <input type="text" id="newLandIdInput" placeholder="例如：8888" autofocus 
                               style="font-size: 15px; padding: 10px;">
                    </div>
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label class="field-label" style="font-size: 14px; margin-bottom: 8px;"><span style="color:#d93025;">*</span> 土地總面積 (m²)</label>
                        <input type="number" id="newLandAreaInput" placeholder="例如：100" 
                               style="font-size: 15px; padding: 10px;" step="0.01">
                    </div>
                    <div class="form-group">
                        <label class="field-label" style="font-size: 14px; margin-bottom: 8px;"><span style="color:#d93025;">*</span> 所有權人姓名</label>
                        <input type="text" id="newOwnerNameInput" placeholder="例如：王小明" 
                               style="font-size: 15px; padding: 10px;">
                    </div>
                    <p id="newNumError" style="color: #d93025; font-size: 13px; margin-top: 15px; display: none;"></p>
                </div>
                <div class="modal-footer">
                    <button class="btn-cancel" onclick="closeNewNumModal()">取消</button>
                    <button class="btn-save" onclick="confirmNewNum()">確定</button>
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHTML);

    setTimeout(() => {
        const input = document.getElementById('newLandIdInput');
        if(input) {
            input.focus();
            // 綁定 Enter 鍵觸發確認
            ['newLandIdInput', 'newLandAreaInput', 'newOwnerNameInput'].forEach(id => {
                document.getElementById(id).addEventListener('keypress', function (e) {
                    if (e.key === 'Enter') confirmNewNum();
                });
            });
        }
    }, 100);
}

function confirmNewNum() {
    const idInput = document.getElementById('newLandIdInput');
    const areaInput = document.getElementById('newLandAreaInput');
    const nameInput = document.getElementById('newOwnerNameInput');
    const errorMsg = document.getElementById('newNumError');
    
    const newId = idInput.value.trim();
    const area = parseFloat(areaInput.value) || 0;
    const ownerName = nameInput.value.trim();

    // 先恢復原本的邊框顏色
    [idInput, areaInput, nameInput].forEach(el => el.style.borderColor = '#ddd');

    // 防呆驗證
    if (!newId || area <= 0 || !ownerName) {
        errorMsg.innerText = "請填寫所有必填欄位 (*)，且面積須大於 0。";
        errorMsg.style.display = 'block';
        if (!newId) idInput.style.borderColor = '#d93025';
        if (area <= 0) areaInput.style.borderColor = '#d93025';
        if (!ownerName) nameInput.style.borderColor = '#d93025';
        return;
    }

    if (data[newId]) {
        errorMsg.innerText = "此地號已存在，請勿重複新增。";
        errorMsg.style.display = 'block';
        idInput.style.borderColor = '#d93025';
        return;
    }

    // 初始化結構並直接帶入使用者剛剛輸入的基礎資料
    data[newId] = {
        '所有權人': [ownerName],
        '土地面積': { '面積': area, '權利範圍': [1] }, // 預設權利範圍為全部(1)
        '他項權利': '',
        '公告現值': [''],
        '年月': [''],
        '當期公告現值': 0,
        '增值稅預估(自用)': [''],
        '增值稅預估(一般)': [''],
        '建號': [''], '建物門牌': [''], '樓層': [''], '建物面積': [''], '權利範圍': [''], '所有權地址': [''], '電話': '', '增值稅試算(自用)': [''], '增值稅試算(一般)': [''],
        '使用分區': "商", '基準容積率': 4.4, '建蔽率': 0.7
    };
    
    if(typeof window.saveProjectToLocal === 'function') window.saveProjectToLocal();
    closeNewNumModal();
    
    // 刷新頁面並開啟剛新增的編輯視窗
    const activeTab = document.querySelector('.nav-item.active').id;
    show(activeTab); 
    setTimeout(() => { openEditModal(newId); }, 300);
}

function closeNewNumModal() {
    const modal = document.getElementById('newNumModal');
    if (modal) modal.remove();
}

// 開啟 Modal
function openEditModal(id) {
    currentEditingId = id;
    const record = data[id];
    if(!record) { newNum(); return; }

    document.getElementById('modalTitle').innerText = id;

    const totalArea = record['土地面積']['面積'];
    document.getElementById('editTotalArea').value = totalArea;
    document.getElementById('editTotalPing').value = (totalArea * PING_FACTOR).toFixed(2);
    document.getElementById('editCurrentValue102').value = record['當期公告現值'] || '';
    document.getElementById('editOtherRights').value = record['他項權利'] || '';

    const listContainer = document.getElementById('ownerList');
    listContainer.innerHTML = ''; 

    const owners = record['所有權人'];
    const scopes = record['土地面積']['權利範圍'] || [];
    const dates = record['年月'] || [];
    const values = record['公告現值'] || [];
    const taxSelf = record['增值稅預估(自用)'] || [];
    const taxGen = record['增值稅預估(一般)'] || [];
    
    // 建物資料
    const buildNos = record['建號'] || [];
    const buildAddrs = record['建物門牌'] || [];
    const floors = record['樓層'] || [];
    const buildAreas = record['建物面積'] || [];
    const buildScopes = record['權利範圍'] || []; 
    const ownerAddrs = record['所有權地址'] || [];
    const calcTaxSelf = record['增值稅試算(自用)'] || [];
    const calcTaxGen = record['增值稅試算(一般)'] || [];

    for(let i=0; i < owners.length; i++) {
        let scopeVal = Array.isArray(scopes) ? scopes[i] : scopes;
        let ownerData = {
            name: owners[i],
            scope: scopeVal,
            date: dates[i] || '',
            value: values[i] || '',
            taxSelf: taxSelf[i] || '',
            taxGen: taxGen[i] || '',
            buildNo: buildNos[i] || '',
            buildAddr: buildAddrs[i] || '',
            floor: floors[i] || '',
            buildArea: buildAreas[i] || '',
            buildScope: buildScopes[i] || '',
            ownerAddr: ownerAddrs[i] || '',
            calcTaxSelf: calcTaxSelf[i] || '',
            calcTaxGen: calcTaxGen[i] || ''
        };
        addOwnerRow(ownerData);
    }
    document.getElementById('editModal').style.display = 'flex';
}

function closeModal() {
    document.getElementById('editModal').style.display = 'none';
    currentEditingId = null;
    // 如果在 B 頁面，刷新一下顯示 (更新總計)
    if(document.getElementById('B').classList.contains('active')) {
        updateGlobalSummary();
    }
    // 如果在 C 或 D，需要刷新卡片
    if(document.getElementById('C').classList.contains('active')) show('C');
    if(document.getElementById('D').classList.contains('active')) show('D');
}

function addOwnerRow(data = {}, isNew = false) {
    const div = document.createElement('div');
    div.className = 'owner-card';
    
    // 解構資料
    const { name='', scope='', date='', value='', taxSelf='', taxGen='', buildNo='', buildAddr='', floor='', buildArea='', buildScope='', ownerAddr='', calcTaxSelf='', calcTaxGen='' } = data;

    div.innerHTML = `
        <div class="owner-card-header">
            <strong style="color:#555;">所有權人資訊</strong>
            <button class="btn-remove-card" onclick="this.closest('.owner-card').remove(); calculateGlobalLayout();">
                <i class="fa-solid fa-trash"></i> 移除
            </button>
        </div>
        
        <div class="grid-2-col" style="margin-bottom:15px;">
            <div><span class="field-label">姓名</span><input type="text" class="input-name" value="${name}" placeholder="請輸入姓名"></div>
            <div><span class="field-label">權利範圍 (土地)</span><input type="number" class="input-scope" value="${scope}" step="0.0001" oninput="calculateRow(this)" placeholder="持分比例"></div>
        </div>

        <div class="grid-3-col" style="margin-bottom:15px; background:#f8f9fa; padding:10px; border-radius:6px; border:1px solid #eee;">
            <div><span class="field-label">土地持分 (m²)</span><input type="text" class="input-held-area" readonly style="background:transparent; border:none; padding:0; font-weight:800; color:#333;"></div>
            <div><span class="field-label">土地持分 (坪)</span><input type="text" class="input-held-ping" readonly style="background:transparent; border:none; padding:0; font-weight:800; color:#333;"></div>
            <div><span class="field-label">總土地比率 (%)</span><input type="text" class="input-held-ratio" readonly style="background:transparent; border:none; padding:0; font-weight:800; color:#333;"></div>
        </div>

        <div class="grid-2-col" style="margin-bottom:15px;">
            <div><span class="field-label">年月</span><input type="text" class="input-date" value="${date}"></div>
            <div><span class="field-label">公告現值</span><input type="number" class="input-value" value="${value}"></div>
        </div>

        <div class="grid-2-col" style="margin-bottom:24px; border-bottom:1px dashed #ddd; padding-bottom:15px;">
            <div><span class="field-label">增值稅預估(自用)</span><input type="number" class="input-tax-self" value="${taxSelf}"></div>
             <div><span class="field-label">增值稅預估(一般)</span><input type="number" class="input-tax-gen" value="${taxGen}"></div>
        </div>

        <h4 style="font-size:13px; color:var(--accent-purple); margin-bottom:10px; border-left:3px solid var(--accent-purple); padding-left:8px;">建物詳細資料</h4>
        <div class="grid-2-col" style="margin-bottom:15px;">
            <div><span class="field-label">建號</span><input type="text" class="input-build-no" value="${buildNo}"></div>
             <div><span class="field-label">建物門牌</span><input type="text" class="input-build-addr" value="${buildAddr}"></div>
        </div>
        <div class="grid-2-col" style="margin-bottom:15px;">
            <div><span class="field-label">樓層</span><input type="text" class="input-floor" value="${floor}"></div>
             <div><span class="field-label">建物總面積 (m²)</span><input type="number" class="input-build-area" value="${buildArea}" oninput="calculateBuildingRow(this)"></div>
        </div>
        <div class="grid-2-col" style="margin-bottom:15px;">
            <div><span class="field-label">建物權利範圍</span><input type="number" class="input-build-scope" value="${buildScope}" step="0.0001" oninput="calculateBuildingRow(this)"></div>
             <div><span class="field-label">所有權地址</span><input type="text" class="input-owner-addr" value="${ownerAddr}"></div>
        </div>
        <div class="grid-3-col" style="margin-bottom:15px; background:#f8f9fa; padding:10px; border-radius:6px; border:1px solid #eee;">
            <div><span class="field-label">建物總坪數</span><input type="text" class="readonly-build-total-ping" readonly style="background:transparent; border:none; padding:0; font-weight:800; color:#333;"></div>
            <div><span class="field-label">建物持分 (m²)</span><input type="text" class="readonly-build-held-area" readonly style="background:transparent; border:none; padding:0; font-weight:800; color:#333;"></div>
            <div><span class="field-label">建物持分 (坪)</span><input type="text" class="readonly-build-held-ping" readonly style="background:transparent; border:none; padding:0; font-weight:800; color:#333;"></div>
        </div>
        <div class="grid-2-col" style="margin-bottom:15px;">
            <div><span class="field-label">增值稅試算(自用)</span><input type="number" class="input-calc-tax-self" value="${calcTaxSelf}"></div>
             <div><span class="field-label">增值稅試算(一般)</span><input type="number" class="input-calc-tax-gen" value="${calcTaxGen}"></div>
        </div>
    `;

    document.getElementById('ownerList').appendChild(div);
    calculateRow(div.querySelector('.input-scope'));      
    calculateBuildingRow(div.querySelector('.input-build-area'));

    if (isNew) {
        requestAnimationFrame(() => {
            div.scrollIntoView({ behavior: 'smooth', block: 'center' });
        });
    }
}

function calculateGlobalLayout() {
    const totalArea = parseFloat(document.getElementById('editTotalArea').value) || 0;
    document.getElementById('editTotalPing').value = (totalArea * PING_FACTOR).toFixed(2);
    const rows = document.querySelectorAll('.owner-card');
    rows.forEach(row => {
        calculateRow(row.querySelector('.input-scope'));
    });
}

function calculateRow(scopeInput) {
    if(!scopeInput) return;
    const card = scopeInput.closest('.owner-card');
    const totalArea = parseFloat(document.getElementById('editTotalArea').value) || 0;
    const scope = parseFloat(scopeInput.value) || 0;

    const heldArea = totalArea * scope;
    const heldPing = heldArea * PING_FACTOR;
    // 簡單防呆總持分為0的情況
    const sumHeld = parseFloat(document.getElementById('sumHeldArea').innerText) || 1; 
    const heldRatio = heldArea / sumHeld; // 這邊的計算邏輯可能需要後端資料支援，暫時以個別計算為主

    card.querySelector('.input-held-area').value = heldArea.toFixed(2);
    card.querySelector('.input-held-ping').value = heldPing.toFixed(2);
    // card.querySelector('.input-held-ratio').value = (heldRatio * 100).toFixed(2); // 暫時移除，因 Modal 內無法即時取得全域總和
}

function calculateBuildingRow(input) {
    if(!input) return;
    const card = input.closest('.owner-card');
    const buildArea = parseFloat(card.querySelector('.input-build-area').value) || 0;
    const buildScope = parseFloat(card.querySelector('.input-build-scope').value) || 0;

    const totalPing = buildArea * PING_FACTOR;
    const heldArea = buildArea * buildScope;
    const heldPing = heldArea * PING_FACTOR;

    card.querySelector('.readonly-build-total-ping').value = totalPing.toFixed(2);
    card.querySelector('.readonly-build-held-area').value = heldArea.toFixed(2);
    card.querySelector('.readonly-build-held-ping').value = heldPing.toFixed(2);
}

function saveEdit() {
    if(!currentEditingId) return;
    const record = data[currentEditingId];

    const newTotalArea = parseFloat(document.getElementById('editTotalArea').value) || 0;
    
    // --- 新增防呆：面積不能為 0 ---
    if (newTotalArea <= 0) {
        alert("請輸入有效的「土地總面積」(需大於 0) 後再儲存！");
        document.getElementById('editTotalArea').focus();
        return;
    }

    const newVal102 = parseFloat(document.getElementById('editCurrentValue102').value) || 0;
    const newOtherRights = document.getElementById('editOtherRights').value;

    const rows = document.querySelectorAll('.owner-card');
    
    let newOwners = [], newScopes = [], newDates = [], newValues = [], newTaxSelf = [], newTaxGen = [];
    let newBuildNos = [], newBuildAddrs = [], newFloors = [], newBuildAreas = [], newBuildScopes = [], newOwnerAddrs = [], newCalcTaxSelf = [], newCalcTaxGen = [];

    rows.forEach(row => {
        const name = row.querySelector('.input-name').value.trim();
        if(name) {
            newOwners.push(name);
            newScopes.push(parseFloat(row.querySelector('.input-scope').value) || 0);
            newDates.push(row.querySelector('.input-date').value);
            newValues.push(parseInt(row.querySelector('.input-value').value) || 0);
            newTaxSelf.push(parseInt(row.querySelector('.input-tax-self').value) || 0);
            newTaxGen.push(parseInt(row.querySelector('.input-tax-gen').value) || 0);
            
            newBuildNos.push(row.querySelector('.input-build-no').value);
            newBuildAddrs.push(row.querySelector('.input-build-addr').value);
            newFloors.push(row.querySelector('.input-floor').value);
            newBuildAreas.push(parseFloat(row.querySelector('.input-build-area').value) || 0);
            newBuildScopes.push(parseFloat(row.querySelector('.input-build-scope').value) || 0);
            newOwnerAddrs.push(row.querySelector('.input-owner-addr').value);
            newCalcTaxSelf.push(parseInt(row.querySelector('.input-calc-tax-self').value) || 0);
            newCalcTaxGen.push(parseInt(row.querySelector('.input-calc-tax-gen').value) || 0);
        }
    });

    // --- 新增防呆：至少要有一個所有權人 ---
    if(newOwners.length === 0) { 
        alert("請至少輸入一位所有權人的「姓名」！"); 
        return; 
    }

    // 更新物件
    record['所有權人'] = newOwners;
    record['土地面積']['面積'] = newTotalArea;
    record['土地面積']['權利範圍'] = newScopes;
    record['年月'] = newDates;
    record['公告現值'] = newValues;
    record['當期公告現值'] = newVal102;
    record['他項權利'] = newOtherRights;
    record['增值稅預估(自用)'] = newTaxSelf;
    record['增值稅預估(一般)'] = newTaxGen;

    record['建號'] = newBuildNos;
    record['建物門牌'] = newBuildAddrs;
    record['樓層'] = newFloors;
    record['建物面積'] = newBuildAreas;
    record['權利範圍'] = newBuildScopes;
    record['所有權地址'] = newOwnerAddrs;
    record['增值稅試算(自用)'] = newCalcTaxSelf;
    record['增值稅試算(一般)'] = newCalcTaxGen;

    // 如果圖表存在則更新
    const chartId = `chart_${currentEditingId}`;
    if(charts[chartId]) {
        const targetChart = charts[chartId];
        targetChart.data.labels = [];
        targetChart.data.datasets[0].data = [];
        targetChart.data.datasets[0].backgroundColor = [];

        for(let i=0; i<newOwners.length; i++) {
            targetChart.data.labels.push(newOwners[i]);
            targetChart.data.datasets[0].data.push(newTotalArea * newScopes[i]);
            targetChart.data.datasets[0].backgroundColor.push(colorPalette[i % colorPalette.length]);
        }
        targetChart.update();
        updateCenterTotal(chartId);
    }
    if(typeof window.saveProjectToLocal === 'function') window.saveProjectToLocal();
    closeModal();
}

function updateGlobalSummary() {
    let totalHeldM2 = 0;
    let totalValue = 0;

    for (let id in data) {
        if(id === '地段') continue;
        let land = data[id];
        let landArea = land['土地面積']['面積'] || 0;
        let landUnitValue = land['當期公告現值'] || 0;
        let scopes = land['土地面積']['權利範圍'] || [];
        
        let currentLandScopeSum = 0;
        // 注意：這裡應該要加總所有權人的持分，通常總和是 1，但為了保險起見我們加總
        let owners = land['所有權人'] || [];
        for (let i = 0; i < owners.length; i++) {
             let s = Array.isArray(scopes) ? scopes[i] : scopes;
             currentLandScopeSum += (parseFloat(s) || 0);
        }

        totalHeldM2 += landArea * currentLandScopeSum;
        totalValue += landUnitValue * currentLandScopeSum; 
    }

    const sumAreaEl = document.getElementById('sumHeldArea');
    const sumPingEl = document.getElementById('sumHeldPing');
    const sumValueEl = document.getElementById('sumAssessedValue');

    if (sumAreaEl) sumAreaEl.innerText = totalHeldM2.toFixed(2);
    if (sumPingEl) sumPingEl.innerText = (totalHeldM2 * PING_FACTOR).toFixed(2);
    if (sumValueEl) sumValueEl.innerText = '$' + Math.round(totalValue).toLocaleString();
}



// 1. 處理單一 PDF 檔案並轉為文字
async function processTranscriptPDF(file) {
    try {
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument(arrayBuffer);
        const pdf = await loadingTask.promise;
        let fullText = "";

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            fullText += textContent.items.map(item => item.str).join(' ') + "\n";
        }

        parseTranscriptText(fullText);
        if(typeof window.saveProjectToLocal === 'function') window.saveProjectToLocal();
    } catch (error) {
        console.error("解析 PDF 失敗:", error);
        alert(`檔案 ${file.name} 解析失敗，請確認是否為標準謄本。`);
    }
}

// 2. 核心解析函式：將謄本文字轉化為資料庫物件
function parseTranscriptText(text) {
    // A. 預處理：將全形轉半形，並消滅破壞排版的干擾空格
    function normalize(str) {
        return str.replace(/[\uFF01-\uFF5E]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xfee0))
                  .replace(/　/g, " ")
                  .replace(/面\s+積/g, "面積")
                  .replace(/住\s+址/g, "住址")
                  .replace(/層\s+次/g, "層次")
                  .replace(/權\s+利\s+人/g, "權利人");
    }

    const cleanText = normalize(text);

    // B. 以謄本標題作為切分點，處理一份 PDF 內含多筆地號/建號的狀況
    const blocks = cleanText.split(/(?=土地登記第二類謄本|建物登記第二類謄本)/);

    blocks.forEach(block => {
        if (block.includes('土地登記第二類謄本')) {
            parseLandBlock(block);
        } else if (block.includes('建物登記第二類謄本')) {
            parseBuildingBlock(block);
        }
    });
}

// --- 解析土地謄本 ---
function parseLandBlock(block) {
    // 1. 抓取地號
    const landMatch = block.match(/(\d{4}-\d{4})\s*地號/);
    if (!landMatch) return;
    const landId = landMatch[1];

    // 初始化你的 data.js 完整結構
    if (!data[landId]) {
        data[landId] = {
            "所有權人": [],
            "土地面積": { "面積": 0, "權利範圍": [] },
            "他項權利": "",
            "公告現值": [], "年月": [], "當期公告現值": 0,
            "增值稅預估(自用)": [], "增值稅預估(一般)": [],
            "建號": [], "建物門牌": [], "樓層": [], "建物面積": [], "權利範圍": [], "所有權地址": [],
            "電話": "", "增值稅試算(自用)": [], "增值稅試算(一般)": [],
            "使用分區": "商", "基準容積率": 4.4, "建蔽率": 0.7
        };
    }

    // 2. 抓取面積與當期現值 (使用 [*]* 避開遮蔽星號)
    const areaMatch = block.match(/面積[:：]\s*[*]*([\d\.]+)/);
    if (areaMatch) data[landId]["土地面積"]["面積"] = parseFloat(areaMatch[1]);

    const currValMatch = block.match(/公告土地現值[:：]\s*[*]*([\d,]+)/);
    if (currValMatch) data[landId]["當期公告現值"] = parseInt(currValMatch[1].replace(/,/g, ''), 10);

    // 3. 抓取他項權利 (銀行與金額)
    let otherRights = [];
    const rightMatches = [...block.matchAll(/權利人[:：]\s*([^\s]+)/g)];
    const amountMatches = [...block.matchAll(/擔保債權總金額[:：].*?[*]*([\d,]+)\s*元正/g)];
    
    for (let i = 0; i < Math.min(rightMatches.length, amountMatches.length); i++) {
        otherRights.push(`${rightMatches[i][1]}\n設定抵押權${amountMatches[i][1]}元`);
    }
    if (otherRights.length > 0) data[landId]["他項權利"] = otherRights.join("\n\n");

    // 4. 拆分並循環抓取每位「所有權人」細節
    const ownerSegments = block.split(/\(\d+\)\s*登記次序/);
    ownerSegments.shift(); // 移除第一段非所有權人區塊

    ownerSegments.forEach(seg => {
        // 抓取姓名與統編，組成 "姓名(統編)"
        const nameMatch = seg.match(/所有權人[:：]\s*([^\s]+)/);
        if (!nameMatch) return;
        const name = nameMatch[1]; 
        
        const idMatch = seg.match(/統一編號[:：]\s*([A-Z0-9*]+)/);
        const idStr = idMatch ? idMatch[1] : '';
        const fullName = idStr ? `${name}(${idStr})` : name;

        // 抓取土地權利範圍 (處理分之)
        const scopeMatch = seg.match(/權利範圍[:：]\s*[*]*(\d+)\s*分之\s*(\d+)/);
        let scope = 1.0;
        if (scopeMatch) {
            scope = parseFloat(scopeMatch[2]) / parseFloat(scopeMatch[1]); // 分子 / 分母
        } else if (seg.match(/權利範圍[:：]\s*[*]*全部/)) {
            scope = 1.0;
        }

        // 抓取住址
        const addrMatch = seg.match(/住址[:：]\s*([^\n\r(（]+)/);
        const address = addrMatch ? addrMatch[1].trim() : '';

        // 抓取前次移轉現值與年月 (例如: 108年08月 **123,000.0)
        const prevMatch = seg.match(/前次移轉現值[\s\S]*?(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*[*]*([\d,.]+)\s*元/);
        let prevDate = '', prevVal = 0;
        if (prevMatch) {
            prevDate = `${prevMatch[1]}/${prevMatch[2]}`; // 組成 "108/08"
            prevVal = parseFloat(prevMatch[3].replace(/,/g, ''));
        }

        // 確保該所有權人被加入，並將陣列推齊
        let idx = data[landId]["所有權人"].indexOf(fullName);
        if (idx === -1) {
            data[landId]["所有權人"].push(fullName);
            data[landId]["土地面積"]["權利範圍"].push(scope);
            data[landId]["所有權地址"].push(address);
            data[landId]["年月"].push(prevDate);
            data[landId]["公告現值"].push(prevVal);
            
            // 其餘需要保持陣列長度一致的欄位先塞空值
            data[landId]["增值稅預估(自用)"].push("");
            data[landId]["增值稅預估(一般)"].push("");
            data[landId]["建號"].push("");
            data[landId]["建物門牌"].push("");
            data[landId]["樓層"].push("");
            data[landId]["建物面積"].push("");
            data[landId]["權利範圍"].push("");
            data[landId]["增值稅試算(自用)"].push("");
            data[landId]["增值稅試算(一般)"].push("");
        }
    });
}

// --- 解析建物謄本 ---
function parseBuildingBlock(block) {
    const buildMatch = block.match(/(\d{5}-\d{3})\s*建號/);
    if (!buildMatch) return;
    const buildId = buildMatch[1];

    const landMatch = block.match(/建物坐落地號[:：].*?(\d{4}-\d{4})/);
    if (!landMatch) return;
    const landId = landMatch[1];

    // 如果建物找不到對應的土地資料，就先忽略
    if (!data[landId]) return;

    const addrMatch = block.match(/建物門牌[:：]\s*([^\n\r]+)/);
    const address = addrMatch ? addrMatch[1].trim() : '';

    const areaMatch = block.match(/總面積[:：]\s*[*]*([\d\.]+)/);
    const buildArea = areaMatch ? parseFloat(areaMatch[1]) : 0;

    const floorMatch = block.match(/層次[:：]\s*([^\s]+)/);
    const floor = floorMatch ? floorMatch[1].trim() : '';

    // 處理建物所有權人，將建物數據對應寫回土地的陣列中
    const ownerSegments = block.split(/\(\d+\)\s*登記次序/);
    ownerSegments.shift();

    ownerSegments.forEach(seg => {
        const nameMatch = seg.match(/所有權人[:：]\s*([^\s]+)/);
        if (!nameMatch) return;
        const name = nameMatch[1]; 

        const idMatch = seg.match(/統一編號[:：]\s*([A-Z0-9*]+)/);
        const idStr = idMatch ? idMatch[1] : '';
        const fullName = idStr ? `${name}(${idStr})` : name;

        const scopeMatch = seg.match(/權利範圍[:：]\s*[*]*(\d+)\s*分之\s*(\d+)/);
        let scope = 1.0;
        if (scopeMatch) {
            scope = parseFloat(scopeMatch[2]) / parseFloat(scopeMatch[1]);
        } else if (seg.match(/權利範圍[:：]\s*[*]*全部/)) {
            scope = 1.0;
        }

        // 透過完整的 "姓名(統編)" 尋找在地號中的索引，填入建物資料
        let idx = data[landId]["所有權人"].indexOf(fullName);
        if (idx !== -1) {
            data[landId]["建號"][idx] = buildId;
            data[landId]["建物門牌"][idx] = address;
            data[landId]["樓層"][idx] = floor;
            data[landId]["建物面積"][idx] = buildArea;
            data[landId]["權利範圍"][idx] = scope; // 覆蓋為建物的持分比例
        }
    });

}

