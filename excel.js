const alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB'];
async function exportFancyExcel(shift) {
    const workbook = new ExcelJS.Workbook();
    
    // 建立五個工作表
    const sheetA = workbook.addWorksheet('清冊');
    const sheetB = workbook.addWorksheet('增值稅歸戶表');
    const sheetC = workbook.addWorksheet('土地歸戶表');
    const sheetD = workbook.addWorksheet('地主可分配面積');
    const sheetE = workbook.addWorksheet('基本分析'); // 修正名稱
    let total = 0;
    for(let k in data){
        total += data[k]['土地面積']['面積'];
    }
    // -------------------------------------------------------
    // 1. Sheet A: 清冊
    // -------------------------------------------------------
    function goA(){
        sheetA.columns = [
            { key: '地段', width: 10 },
            { key: '小段', width: 10 },
            { key: '地號', width: 10 },
            { key: '所有權人', width: 15 },
            // 土地面積群組
            { key: '土地面積_m²', width: 12 , style: { numFmt: '0.00' }},
            { key: '土地面積_坪', width: 12 , style: { numFmt: '0.00' }},
            { key: '土地面積_權利範圍', width: 12 , style: { numFmt: '# ?/?' }},
            { key: '土地面積_土地持分面積(m²)', width: 19 , style: { numFmt: '0.00' }},
            { key: '土地面積_土地持分面積(坪)', width: 19 , style: { numFmt: '0.00' }},
            { key: '土地面積_總土地比率', width: 13, style: { numFmt: '0.00%' }},
            // 其他資訊
            { key: '他項權利', width: 15 },
            { key: '前次取得_公告現值', width: 15 , style: { numFmt: '#,##0.00' }},
            { key: '前次取得_年月', width: 12 },
            { key: '當期公告現值', width: 22  , style: { numFmt: '#,##0.00' }},
            { key: '增值稅預估(自用)', width: 18 , style: { numFmt: '#,##0.00' }},
            { key: '增值稅預估(一般)', width: 18 , style: { numFmt: '#,##0.00' }},
            // 建物資訊
            { key: '建號', width: 10 },
            { key: '建物門牌', width: 25 },
            { key: '樓層', width: 10 },
            { key: '建物面積_m²', width: 12 , style: { numFmt: '0.00' }},
            { key: '建物面積_坪', width: 12 , style: { numFmt: '0.00' }},
            { key: '權利範圍', width: 12 , style: { numFmt: '# ?/?' }},
            { key: '持分面積_m²', width: 12 , style: { numFmt: '0.00' }},
            { key: '持分面積_坪', width: 12 , style: { numFmt: '0.00' }},
            { key: '所有權地址', width: 30 },
            { key: '電話', width: 15 },
            { key: '增值稅試算_一般', width: 15 , style: { numFmt: '#,##0.00' }},
            { key: '增值稅試算_自用', width: 15 , style: { numFmt: '#,##0.00' }},
        ];
        if (sheetA.columnCount > 29) {
            sheetA.spliceColumns(30, sheetA.columnCount - 29);
        }
        sheetA.addRow([`${basics['地段']}土地清冊`]);
        let sheetArow2 = ['地段','小段','地號','所有權人','土地面積','','','','','','他項權利','前次取得','','當期公告現值','增值稅預估(自用)','增值稅預估(一般)','建號','建物門牌','樓層','建物面積','','權利範圍','持分面積','','所有權地址','電話','增值稅試算',''];
        let sheetArow3 = ['','','','','m²','坪','權利範圍','土地持分面積(m²)','土地持分面積(坪)','總土地比率','','公告現值','年月','','','','','','','m²','坪','','m²','坪','','','一般','自用'];
        sheetA.addRow(sheetArow2);
        sheetA.addRow(sheetArow3);
        sheetA.mergeCells('A1:AB1');
        let first = 0;
        for(let i = 0 ; i<28 ; i++){
            if(sheetArow3[i] === ''){
                sheetA.mergeCells(`${alphabet[i]}2:${alphabet[i]}3`);
            }
            if(sheetArow2[i] === '' && first === 0){
                first = i-1;
            }
            if((sheetArow2[i] !== '' || i===27) && first!==0){
                if(i===27){
                    sheetA.mergeCells(`${alphabet[first]}2:AB2`);
                }
                else{
                    sheetA.mergeCells(`${alphabet[first]}2:${alphabet[i-1]}2`);
                }
                first = 0;
            }

            [`${alphabet[i]}2`,`${alphabet[i]}3`].forEach(cellKey => {
                const cell = sheetA.getCell(cellKey);
                
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFd6e3bc' }
                };

                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });

        }
        for(let num in data){
            let temp = [];
            const area = data[num]['土地面積']['面積'];
            for(let i = 0 ; i<data[num]['所有權人'].length ; i++){
                const ratio = data[num]['土地面積']['權利範圍'][i];
                if(shift){
                    temp.push(['','','',data[num]['所有權人'][i],'','',ratio,area*ratio,area*ratio*0.3025,area*ratio/total,'',data[num]['公告現值'][i],data[num]['年月'][i],'',data[num]['增值稅預估(自用)'][i],data[num]['增值稅預估(一般)'][i],data[num]['建號'][i],data[num]['建物門牌'][i],data[num]['樓層'][i],data[num]['建物面積'][i],data[num]['建物面積'][i]*0.3025,data[num]['權利範圍'][i],data[num]['建物面積'][i]*data[num]['權利範圍'][i],data[num]['建物面積'][i]*data[num]['權利範圍'][i]*0.3025,data[num]['所有權地址'][i],'',data[num]['增值稅試算(自用)'][i],data[num]['增值稅試算(一般)'][i]]);
                }
                else{
                    temp.push(['','','',data[num]['所有權人'][i],'','',ratio,{formula:`E${sheetA.rowCount+1}*G${sheetA.rowCount+1+i}`},{formula:`H${sheetA.rowCount+1+i}*0.3025`},{formula:`H${sheetA.rowCount+1+i}/SUM(H4:H9999)*2`},'',data[num]['公告現值'][i],data[num]['年月'][i],'',data[num]['增值稅預估(自用)'][i],data[num]['增值稅預估(一般)'][i],data[num]['建號'][i],data[num]['建物門牌'][i],data[num]['樓層'][i],data[num]['建物面積'][i],{formula:`T${sheetA.rowCount+1+i}*0.3025`},data[num]['權利範圍'][i],{formula:`T${sheetA.rowCount+1+i}*V${sheetA.rowCount+1+i}`},{formula:`W${sheetA.rowCount+1+i}*0.3025`},data[num]['所有權地址'][i],'',data[num]['增值稅試算(自用)'][i],data[num]['增值稅試算(一般)'][i]]);
                }
            }
            temp[0][2] = num;
            temp[0][4] = area;
            if(shift){
                temp[0][5] = area*0.3025;
            }
            else{
                temp[0][5] = { formula: `E${sheetA.rowCount+1}*0.3025` };
            }
            
            temp[0][10] = data[num]['他項權利'];
            temp[0][13] = data[num]['當期公告現值'];
            temp[0][25] = data[num]['電話'];

            let last = sheetA.rowCount+1;
            for(let k of temp){
                sheetA.addRow(k);
            }

            ['C','E','F','K','N','Z'].forEach(key => {
                sheetA.mergeCells(`${key}${last}:${key}${sheetA.rowCount}`);
            });

            ['L','M'].forEach(key => {
                mergeCheck(sheetA,key,last,sheetA.rowCount);
            });
        }
        
        ['D','Q','R','Y'].forEach(key => {
            mergeCheck(sheetA,key,4,sheetA.rowCount);
        });

        sheetA.mergeCells(`A4:A${sheetA.rowCount}`);
        
        // 1. 準備資料與新增列
        let finalRow = ['','','','總計'];
        for(let i = 4 ; i<28 ; i++){
            finalRow.push('');
        }

        // 新增列
        sheetA.addRow(finalRow);

        // 2. 取得剛剛新增的那一列 (Total Row)
        let currentRowIndex = sheetA.rowCount;
        let totalRow = sheetA.getRow(currentRowIndex);

        // 3. 設定樣式 (灰色背景)
        // 使用 { includeEmpty: true } 確保空白格子也會被上色
        totalRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFf2f2f2' } // 你的指定顏色
            };
            
            // (選用) 通常總計列也會加粗字體，若需要可取消註解
            // cell.font = { bold: true };
            // cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        });

        // 4. 合併儲存格
        sheetA.mergeCells(`A${currentRowIndex}:C${currentRowIndex}`);
        sheetA.mergeCells(`K${currentRowIndex}:M${currentRowIndex}`);
        sheetA.mergeCells(`O${currentRowIndex}:AB${currentRowIndex}`);

        // 5. 寫入 SUM 公式
        ['E','F','H','I','J','N'].forEach(key => {
            // 資料範圍結束於總計列的「上一列」
            const dataEndRow = currentRowIndex - 1; 

            // 將公式寫入當前列
            // 注意：用 value = { formula: ... } 是比較標準的寫法
            if(shift){
                let ans = 0;
                for(let j = 4 ; j<=dataEndRow ; j++){
                    ans += parseFloat(sheetA.getCell(`${key}${j}`).value);
                }
                sheetA.getCell(`${key}${currentRowIndex}`).value = ans;
            }
            else{
                sheetA.getCell(`${key}${currentRowIndex}`).value = {
                    formula: `SUM(${key}4:${key}${dataEndRow})`
                };
            }
            
        });
        

        for (let i = 1; i <= 28; i++) {
            const col = sheetA.getColumn(i);
            col.alignment = { 
                vertical: 'middle', 
                horizontal: col.alignment ? col.alignment.horizontal : undefined, // 保留原本的水平設定
                wrapText: true
            };
        }
        sheetA.getCell('A4').value = basics['地段'];
    }
    goA();

    // -------------------------------------------------------
    // 2. Sheet B: 增值稅歸戶表
    // -------------------------------------------------------
    function goB(){
        sheetB.columns = [
            { key: '編號', width: 8 },
            { key: '所有權人', width: 15 },
            { key: '地號', width: 10 },
            { key: '面積(坪)', width: 12 , style: { numFmt: '0.00' }},
            { key: '權利範圍', width: 12 , style: { numFmt: '# ?/?' }},
            { key: '各持分面積(坪)', width: 16 , style: { numFmt: '0.00' }},
            { key: '前次取得-年月', width: 12 },
            { key: '前次取得-公告現值', width: 15 , style: { numFmt: '#,##0.00' }},
            { key: '當期公告現值', width: 22 , style: { numFmt: '#,##0.00' }},
            { key: '增值稅預估(自用)', width: 19 , style: { numFmt: '#,##0.00' }},
            { key: '增值稅合計(自用)', width: 19 , style: { numFmt: '#,##0.00' }},
            { key: '增值稅預估(一般)', width: 19 , style: { numFmt: '#,##0.00' }},
            { key: '增值稅合計(一般)', width: 19 , style: { numFmt: '#,##0.00' }}
        ];

        const totalColsB = sheetB.columnCount;
        if (totalColsB > 13) {
            //sheetB.spliceColumns(14, totalColsB - 13);
            // 從第 14 欄 (N) 開始，刪除後面 1000 欄 (或更多，確保涵蓋到 Z 甚至更遠)
            sheetB.spliceColumns(14, 16384);
        }

        sheetB.addRow([`${basics['地段']}增值稅歸戶表`]);
        let sheetArow2 = ['編號','所有權人','地號','面積(坪)','權利範圍','各持分面積(坪)','前次取得','','當期公告現值','增值稅預估(自用)','增值稅合計(自用)','增值稅預估(一般)','增值稅合計(一般)'];
        let sheetArow3 = ['','','','','','','年月','公告現值','','','','',''];

        sheetB.addRow(sheetArow2);
        sheetB.addRow(sheetArow3);
        sheetB.mergeCells('A1:M1');

        // --- 處理標題合併 ---
        let first = 0; // 修正：確保變數有初始化
        for(let i = 0 ; i<13 ; i++){
            // 垂直合併檢查
            if(sheetArow3[i] === ''){
                sheetB.mergeCells(`${alphabet[i]}2:${alphabet[i]}3`);
            }

            // 水平合併檢查
            if(sheetArow2[i] === '' && first === 0){
                first = i-1;
            }
            if((sheetArow2[i] !== '' || i===12) && first!==0){ // 修正：i===27 改為 12 (最後一欄)
                // 確保 first 沒跑掉
                if (first < 0) first = 0; 
                
                if(i===12 && sheetArow2[i] === ''){ 
                    // 如果最後一欄也是空，合併到最後
                    sheetB.mergeCells(`${alphabet[first]}2:M2`);
                }
                else{
                    sheetB.mergeCells(`${alphabet[first]}2:${alphabet[i-1]}2`);
                }
                first = 0;
            }

            // 樣式設定
            [`${alphabet[i]}2`,`${alphabet[i]}3`].forEach(cellKey => {
                const cell = sheetB.getCell(cellKey);
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFd6e3bc' }
                };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });
        }

        // --- 資料處理 ---
        let names = []; // 建議從空陣列開始，除非你要過濾掉自己
        let temp = [];

        for(let num in data){
            // data[num]['所有權人'] 是一個陣列，用 let i = 0 ...
            const owners = data[num]['所有權人'];
            for(let i = 0; i < owners.length; i++){
                let name = owners[i];
                
                // 修正 2: 使用 includes 檢查重複
                if(!names.includes(name)){
                    names.push(name);
                }

                let area = data[num]['土地面積']['面積'] * 0.3025;
                let ratio = data[num]['土地面積']['權利範圍'][i]; 
                
                // 修正 1: push 使用圓括號 () 並且放入正確的結構
                // 注意：這裡 names.indexOf(name) + 1 是為了讓編號從 1 開始
                temp.push([
                    names.indexOf(name) + 1, // 編號
                    name,                    // 所有權人
                    num,                     // 地號
                    area,                    // 面積
                    ratio,                   // 權利範圍
                    area*ratio, // 稍後替換公式
                    data[num]['年月'][i],
                    data[num]['公告現值'][i],
                    data[num]['當期公告現值'],
                    data[num]['增值稅預估(自用)'][i],
                    '', // 增值稅合計(自用) - 等合併後填
                    data[num]['增值稅預估(一般)'][i],
                    ''  // 增值稅合計(一般) - 等合併後填
                ]);
            }
        }

        // 根據編號排序 temp，確保相同人的資料在一起
        temp.sort((a, b) => a[0] - b[0]);

        // --- 寫入資料與合併 ---
        for(let i = 1 ; i <= names.length ; i++){
            const temp2 = temp.filter(row => row[0] === i);
            
            if (temp2.length === 0) continue; // 防呆

            let last = sheetB.rowCount + 1; // 資料開始的列
            
            // 寫入資料列
            for(let k of temp2){
                // 修正公式列號 (因為現在才知道寫入第幾列)
                let currentRow = sheetB.rowCount + 1;
                // 複製一份 k 以免改到原始資料
                let rowData = [...k]; 
                // 修正 F 欄公式: D*E
                if(!shift){
                    rowData[5] = { formula: `D${currentRow}*E${currentRow}` };
                }
                sheetB.addRow(rowData);
            }

            ['A','B','G','H','I','K','M'].forEach(key=>{
                mergeCheck(sheetB,key,last,sheetB.rowCount);
                if(key === 'K' || key === 'M'){
                    let preKey = 'Tommy';
                    if(key === 'K'){
                        preKey = 'J';
                    }
                    else if(key === 'M'){
                        preKey = 'L';
                    }
                    if(shift){
                        let ans = 0;
                        for(let j = last ; j<=sheetB.rowCount ; j++){
                            ans += sheetB.getCell(`${preKey}${j}`).value;
                        }
                        sheetB.getCell(`${key}${last}`).value = ans;
                    }
                    else{
                        sheetB.getCell(`${key}${last}`).value = {formula:`SUM(${preKey}${last}:${preKey}${sheetB.rowCount})`};
                    }
                }
            });
        }
        
        finalRow = ['合計','','',''];
        for(let i = 4 ; i<13 ; i++){
            finalRow.push('');
        }

        // 新增列
        sheetB.addRow(finalRow);

        // 2. 取得剛剛新增的那一列 (Total Row)
        currentRowIndex = sheetB.rowCount;
        totalRow = sheetB.getRow(currentRowIndex);

        // 3. 設定樣式 (灰色背景)
        // 使用 { includeEmpty: true } 確保空白格子也會被上色
        totalRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFf2f2f2' } // 你的指定顏色
            };
            
            // (選用) 通常總計列也會加粗字體，若需要可取消註解
            // cell.font = { bold: true };
            // cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        });

        // 4. 合併儲存格
        sheetB.mergeCells(`A${currentRowIndex}:D${currentRowIndex}`);
        sheetB.mergeCells(`G${currentRowIndex}:I${currentRowIndex}`);

        // 5. 寫入 SUM 公式
        ['F','J','K','L','M'].forEach(key => {
            // 資料範圍結束於總計列的「上一列」
            const dataEndRow = currentRowIndex - 1; 
            if(shift){
                let ans = 0;
                for(let j = 4 ; j<=dataEndRow ; j++){
                    ans += sheetB.getCell(`${key}${j}`).value;
                }
                sheetB.getCell(`${key}${currentRowIndex}`).value = ans;
            }
            else{
                sheetB.getCell(`${key}${currentRowIndex}`).value = {
                    formula: `SUM(${key}4:${key}${dataEndRow})`
                };
            }
        });

        for (let i = 1; i <= 13; i++) {
            const col = sheetB.getColumn(i);
            col.alignment = { 
                vertical: 'middle', 
                horizontal: col.alignment ? col.alignment.horizontal : undefined, // 保留原本的水平設定
                wrapText: true
            };
        }
    }
    goB();

    // -------------------------------------------------------
    // 3. Sheet C: 土地歸戶表
    // -------------------------------------------------------
    function goC(){
        sheetC.columns = [
            {key: 'id', width: 8 },
            {key: 'owner', width: 15 },
            {key: 'land_no', width: 10 },
            {key: 'area_ping', width: 12 , style: { numFmt: '0.00' }},
            {key: 'scope', width: 12 , style: { numFmt: '# ?/?' }},
            {key: 'shared_area_ping', width: 18 , style: { numFmt: '0.00' }},
            {key: 'total_shared_area_ping', width: 18 , style: { numFmt: '0.00' }},
            {key: 'est_property_area1', width: 9 , style: { numFmt: '0.00' }},
            {key: 'est_property_area2', width: 9 },
            {key: 'est_property_area3', width: 9 },
            {key: 'joint_allocation', width: 18 , style: { numFmt: '0.00' }}
        ];

        if (sheetC.columnCount > 11) {
            sheetC.spliceColumns(12, sheetC.columnCount - 11);
        }

        sheetC.addRow([`${basics['地段']}增值稅歸戶表`]);
        let sheetArow2 = ['編號','所有權人','地號','面積(坪)','權利範圍','各持分面積(坪)','總持分面積(坪)','預估產權面積(坪)','','','合建分取'];
        let sheetArow3 = ['','','','','','','',basics['土地歸戶表']['預估產權面積(坪)'][0],basics['土地歸戶表']['預估產權面積(坪)'][1],basics['土地歸戶表']['預估產權面積(坪)'][2],basics['土地歸戶表']['合建分取']];
        sheetC.addRow(sheetArow2);
        sheetC.addRow(sheetArow3);
        sheetC.mergeCells('A1:K1');

        sheetC.getCell('K3').style = { numFmt: '0%' };
        let first = 0; // 修正：確保變數有初始化
        for(let i = 0 ; i<11 ; i++){
            // 垂直合併檢查
            if(sheetArow3[i] === ''){
                sheetC.mergeCells(`${alphabet[i]}2:${alphabet[i]}3`);
            }

            // 水平合併檢查
            if(sheetArow2[i] === '' && first === 0){
                first = i-1;
            }
            if((sheetArow2[i] !== '' || i===12) && first!==0){ // 修正：i===27 改為 12 (最後一欄)
                // 確保 first 沒跑掉
                if (first < 0) first = 0; 
                
                if(i===12 && sheetArow2[i] === ''){ 
                    // 如果最後一欄也是空，合併到最後
                    sheetC.mergeCells(`${alphabet[first]}2:M2`);
                }
                else{
                    sheetC.mergeCells(`${alphabet[first]}2:${alphabet[i-1]}2`);
                }
                first = 0;
            }

            // 樣式設定
            [`${alphabet[i]}2`,`${alphabet[i]}3`].forEach(cellKey => {
                const cell = sheetC.getCell(cellKey);
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFd6e3bc' }
                };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });
        }

        let names = [];
        let temp = [];

        for(let num in data){
            // data[num]['所有權人'] 是一個陣列，用 let i = 0 ...
            const owners = data[num]['所有權人'];
            for(let i = 0; i < owners.length; i++){
                let name = owners[i];
                
                // 修正 2: 使用 includes 檢查重複
                if(!names.includes(name)){
                    names.push(name);
                }

                let area = data[num]['土地面積']['面積'] * 0.3025;
                let ratio = data[num]['土地面積']['權利範圍'][i]; 
                
                // 修正 1: push 使用圓括號 () 並且放入正確的結構
                // 注意：這裡 names.indexOf(name) + 1 是為了讓編號從 1 開始
                temp.push([
                    names.indexOf(name) + 1, // 編號
                    name,                    // 所有權人
                    num,                     // 地號
                    area,                    // 面積
                    ratio,                   // 權利範圍
                    area*ratio,
                    '',
                    '',
                    '', // 稍後替換公式
                ]);
            }
        }
        temp.sort((a, b) => a[0] - b[0]);

        // --- 寫入資料與合併 ---
        for(let i = 1 ; i <= names.length ; i++){
            const temp2 = temp.filter(row => row[0] === i);
            
            if (temp2.length === 0) continue; // 防呆

            let last = sheetC.rowCount + 1; // 資料開始的列
            
            // 寫入資料列
            for(let k of temp2){
                // 修正公式列號 (因為現在才知道寫入第幾列)
                let currentRow = sheetC.rowCount + 1;
                // 複製一份 k 以免改到原始資料
                let rowData = [...k]; 
                // 修正 F 欄公式: D*E
                if(!shift){
                    rowData[5] = { formula: `D${currentRow}*E${currentRow}` };
                }
                sheetC.addRow(rowData);
            }

            ['A','B'].forEach(key=>{
                mergeCheck(sheetC,key,last,sheetC.rowCount);
            });
            sheetC.mergeCells(`G${last}:G${sheetC.rowCount}`);
            sheetC.mergeCells(`H${last}:J${sheetC.rowCount}`);
            sheetC.mergeCells(`K${last}:K${sheetC.rowCount}`);
            if(shift){
                let ans = 0;
                for(let j = last ; j<=sheetC.rowCount ; j++){
                    ans += sheetC.getCell(`F${j}`).value;
                }
                sheetC.getCell(`G${last}`).value = ans;
                sheetC.getCell(`H${last}`).value = ans*basics['土地歸戶表']['預估產權面積(坪)'][0]*basics['土地歸戶表']['預估產權面積(坪)'][1]*basics['土地歸戶表']['預估產權面積(坪)'][2];
                sheetC.getCell(`K${last}`).value = ans*basics['土地歸戶表']['預估產權面積(坪)'][0]*basics['土地歸戶表']['預估產權面積(坪)'][1]*basics['土地歸戶表']['預估產權面積(坪)'][2]*total*basics['土地歸戶表']['預估產權面積(坪)'][0]*basics['土地歸戶表']['預估產權面積(坪)'][1]*basics['土地歸戶表']['合建分取'];
            }
            else{
                sheetC.getCell(`G${last}`).value = {formula : `SUM(F${last}:F${sheetC.rowCount})`};
                sheetC.getCell(`H${last}`).value = {formula : `G${last}*H3*I3*J3`};
                sheetC.getCell(`K${last}`).value = {formula : `H${last}*K3`};
            }
        }

        finalRow = ['合計'];
        for(let i = 1 ; i<11 ; i++){
            finalRow.push('');
        }

        // 新增列
        sheetC.addRow(finalRow);

        // 2. 取得剛剛新增的那一列 (Total Row)
        currentRowIndex = sheetC.rowCount;
        totalRow = sheetC.getRow(currentRowIndex);

        // 3. 設定樣式 (灰色背景)
        // 使用 { includeEmpty: true } 確保空白格子也會被上色
        totalRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFf2f2f2' } // 你的指定顏色
            };
            
            // (選用) 通常總計列也會加粗字體，若需要可取消註解
            // cell.font = { bold: true };
            // cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        });

        // 4. 合併儲存格
        sheetC.mergeCells(`A${currentRowIndex}:E${currentRowIndex}`);
        sheetC.mergeCells(`H${currentRowIndex}:J${currentRowIndex}`);

        // 5. 寫入 SUM 公式
        ['F','G','H','K'].forEach(key => {
            const dataEndRow = currentRowIndex - 1;
            if(shift){
                let ans = 0;
                for(let j = 4 ; j<=dataEndRow ; j++){
                    ans += sheetC.getCell(`${key}${j}`).value;
                }
                sheetC.getCell(`${key}${currentRowIndex}`).value = ans;
            }
            else{
                sheetC.getCell(`${key}${currentRowIndex}`).value = {
                    formula: `SUM(${key}4:${key}${dataEndRow})`
                };
            }
        });


        for (let i = 1; i <= 11; i++) {
            const col = sheetC.getColumn(i);
            col.alignment = { 
                vertical: 'middle', 
                horizontal: col.alignment ? col.alignment.horizontal : undefined, // 保留原本的水平設定
                wrapText: true
            };
        }
    }
    goC();

    // -------------------------------------------------------
    // 4. Sheet D: 地主可分配面積 (欄位較多)
    // -------------------------------------------------------
    function goD(){
        sheetD.columns = [
            {key: 'section', width: 10 }, //A
            {key: 'subsection', width: 10 },
            {key: 'land_no', width: 10 },
            {key: 'owner', width: 15 },
            // 土地資料
            {key: 'area_m2', width: 10 , style: { numFmt: '0.00' }},
            {key: 'area_ping', width: 10 , style: { numFmt: '0.00' }},
            {key: 'scope', width: 10 , style: { numFmt: '# ?/?' }},
            {key: 'shared_area_m2', width: 17 , style: { numFmt: '0.00' }},
            {key: 'shared_area_ping', width: 17 , style: { numFmt: '0.00' }},
            {key: 'total_ratio', width: 12 , style: { numFmt: '0.00%' }},
            // 容積計算
            {key: 'zoning', width: 10 }, //K
            {key: 'base_far', width: 12 , style: { numFmt: '0%' }},
            {key: 'coverage', width: 10 , style: { numFmt: '0%' }},
            {key: 'base_far_area', width: 17 , style: { numFmt: '0.00' }},
            {key: 'est_bonus_area', width: 10 , style: { numFmt: '0.00' }},
            {key: 'allowed_total_area', width: 10 , style: { numFmt: '0.00' }},
            // 合建分取
            {key: 'joint_base_alloc', width: 10 , style: { numFmt: '0.00' }}, //Q
            {key: 'joint_bonus_alloc', width: 10 , style: { numFmt: '0.00' }},
            {key: 'est_total_alloc', width: 12 , style: { numFmt: '0.00' }},
            {key: 'est_property_area_ping', width: 12 , style: { numFmt: '0.00%' }},
            {key: 'est_parking', width: 10 , style: { numFmt: '0.00' }},
            // 建物資料 (舊屋)
            {key: 'build_no', width: 10 , style: { numFmt: '0.00' }},
            {key: 'build_owner', width: 10 , style: { numFmt: '0.00' }}, //W
            {key: 'address', width: 10 , style: { numFmt: '0.00' }},
            {key: 'orig_main_m2', width: 10 , style: { numFmt: '0.00' }},
            {key: 'orig_main_ping', width: 10 , style: { numFmt: '0.00' }},
            {key: 'orig_sub_m2', width: 17 , style: { numFmt: '0.00' }}, //AA
            {key: 'orig_sub_ping', width: 15 },
            {key: 'build_scope', width: 10 },
            {key: 'orig_total_m2', width: 15 },
            {key: 'orig_total_ping', width: 15 },
            {key: 'orig_total_ratio', width: 18 },
            {key: 'orig_indoor_ping', width: 18 },
            
            {key: 'post_alloc_area', width: 20 },
            {key: 'post_main_area', width: 18 },
            {key: 'post_public_area', width: 10 },
            {key: 'diff_total', width: 18 },
            {key: 'diff_main', width: 18 },
            {key: 'floor1_exchange', width: 20 },
            {key: 'diff_floor1', width: 20 },
            {key: 'owner_address1', width: 30 },
            {key: 'owner_address2', width: 30 },
            {key: 'owner_address3', width: 30 },
            {key: 'owner_address4', width: 30 },
            {key: 'owner_address5', width: 30 },
            {key: 'owner_address6', width: 30 },
            {key: 'owner_address7', width: 30 },
            {key: 'owner_address8', width: 30 },
        ];

        if (sheetD.columnCount > 48) {
            sheetD.spliceColumns(49, sheetD.columnCount - 48);
        }
        sheetD.addRow([`${basics['地段']}土地清冊`]);
        let sheetArow2 = ['地段','小段','地號','所有權人','土地面積','','','','','','使用分區','基準容積率','建蔽率','基準容積面積(m²)','預估獎勵容積(m²)','','','','允建總容積(m²)','','合建基準容積分取','','合建獎勵容積分取','','預估分取總允建容積','','預估產權面積(坪)','預估車位數','建號','所有權人','門牌','原主建物面積(m²)','原主建物面積(坪)','附屬建物面積(m²)','附屬建物面積(坪)','持分','原建物總面積(m²)','原建物總面積(坪)','原建物總面積比例','原室內面積概算(坪)','合建分配後預估產權面積(坪)','預估分配主建物面積(坪)','預估分配公設面積(坪)','都更前後建物總產權差異(坪)','主建物差異(坪)','1樓換取2樓以上可分面積','1樓建物總產權差異(坪)','所有權人地址'];
        let sheetArow3 = ['','','','','m²','坪','權利範圍','土地持分面積(m²)','土地持分面積(坪)','總土地比率','','','','',basics['地主可分配面積']['預估獎勵容積(m²)'][0],basics['地主可分配面積']['預估獎勵容積(m²)'][1],basics['地主可分配面積']['預估獎勵容積(m²)'][2],basics['地主可分配面積']['預估獎勵容積(m²)'][3],'容積(m²)','總比率(%)',basics['地主可分配面積']['合建基準容積分取'][0],basics['地主可分配面積']['合建基準容積分取'][1],basics['地主可分配面積']['合建獎勵容積分取'][0],basics['地主可分配面積']['合建獎勵容積分取'][1],'m²','坪',basics['地主可分配面積']['預估產權面積(坪)'],basics['地主可分配面積']['預估車位數'],'','','','','','','','','','','','','','','','','','','',''];
        // 先寫入資料
        sheetD.addRow(sheetArow2);
        sheetD.addRow(sheetArow3);
        sheetD.mergeCells('A1:AV1'); // 標題合併

        let mergeStart = -1; // 用來記錄水平合併的起始點 (-1 代表沒有在記錄)

        for (let i = 0; i < 48; i++) {
            // 1. 自動取得欄位字母 (A, B... AA... AV)
            // i 是 0-based，但 getColumn 是 1-based，所以要 +1
            const letter = sheetD.getColumn(i + 1).letter;
            const currentKey = `${letter}2`;
            const nextKey = `${letter}3`;

            // 2. 判斷垂直合併 (上下合併)
            // 條件：第二列有字 且 第三列是空的 (代表這是單一的大標題)
            // 額外條件：mergeStart === -1 (代表目前不在水平合併的過程中，避免衝突)
            if (sheetArow2[i] !== '' && sheetArow3[i] === '' && mergeStart === -1) {
                sheetD.mergeCells(`${currentKey}:${nextKey}`);
            }

            // 3. 判斷水平合併 (左右合併)
            if (sheetArow2[i] === '') {
                // 如果第二列是空的，代表這是某個標題的延伸
                if (mergeStart === -1) {
                    mergeStart = i - 1; // 記錄起始點是「前一欄」
                }
                
                // 如果這是最後一欄 (i=47)，要強制結束合併
                if (i === 47 && mergeStart !== -1) {
                    const startLetter = sheetD.getColumn(mergeStart + 1).letter;
                    sheetD.mergeCells(`${startLetter}2:${letter}2`);
                    mergeStart = -1;
                }
            } else {
                // 如果第二列「有字」，代表新的標題開始了
                // 檢查是否需要結束「上一個」水平合併
                if (mergeStart !== -1) {
                    const startLetter = sheetD.getColumn(mergeStart + 1).letter;
                    const endLetter = sheetD.getColumn(i).letter; // 合併到前一欄 (i-1 + 1)
                    
                    sheetD.mergeCells(`${startLetter}2:${endLetter}2`);
                    mergeStart = -1; // 重置
                }
            }

            // 4. 設定樣式 (針對 A2, A3 ... AV2, AV3)
            [currentKey, nextKey].forEach(cellKey => {
                const cell = sheetD.getCell(cellKey);
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFd6e3bc' }
                };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });
        }
        
        
        const lastRow = sheetA.rowCount;

        // 1. 先填寫公式 (維持你原本的邏輯)
        for (let r = 4; r <= lastRow-1; r++) {
            for (let c = 1; c <= 10; c++) {
                // 取得 SheetA 對應格子的位址 (例如 A4)
                const cellAddress = sheetA.getCell(r, c).address;
                
                // 寫入公式
                if(shift){
                    sheetD.getCell(r, c).value = sheetA.getCell(r, c).value;
                }
                else{
                    sheetD.getCell(r, c).value = {
                        formula: `IF('清冊'!${cellAddress}="","", '清冊'!${cellAddress})`
                    };
                }
                sheetD.getCell(r, c).style = sheetA.getCell(r, c).style;
            }
        }

        // 2. 額外複製「合併儲存格」的狀態 (新增這段)
        // sheetA.model.merges 是一個陣列，裡面存著像 ['A4:A5', 'C4:D4'] 這樣的字串
        if (sheetA.model.merges) {
            sheetA.model.merges.forEach(range => {
                // range 是一個字串，例如 "A4:A5"
                // 我們需要解析它，確認它是否落在我們要複製的範圍內 (Row >= 4 且 Col <= 10)
                
                const [startCell, endCell] = range.split(':');
                
                // 取得該範圍的起始列與結束欄
                const startRow = sheetA.getCell(startCell).row;
                const endCol = sheetA.getCell(endCell).col;

                // 判斷條件：
                // 1. 起始列必須 >= 4 (避開你的標題列)
                // 2. 結束欄必須 <= 10 (只複製 A 到 J 欄的合併，避免影響右邊其他資料)
                if (startRow >= 4 && endCol <= 10) {
                    try {
                        sheetD.mergeCells(range);
                    } catch (e) {
                        console.warn(`無法合併 ${range}，可能與現有合併衝突`, e);
                    }
                }
            });
        }

        sheetD.mergeCells('K2:K3');
        let idx = 3;
        for(let num in data){
            for(let i in data[num]["所有權人"]){
                idx++;
                sheetD.getCell(`K${idx}`).value = data[num]['使用分區'];
                sheetD.getCell(`L${idx}`).value = data[num]['基準容積率'];
                sheetD.getCell(`M${idx}`).value = data[num]['建蔽率'];
            }
        }

        const ratioCell = sheetD.getCell('AA3');
        ratioCell.numFmt = '0.00"倍"';
        
        for(let j = 4 ; j<=sheetD.rowCount ; j++){
            if(shift){
                sheetD.getCell(`N${j}`).value = sheetD.getCell(`H${j}`).value*sheetD.getCell(`L${j}`).value;
                sheetD.getCell(`O${j}`).value = sheetD.getCell(`O3`).value*sheetD.getCell(`N${j}`).value;
                sheetD.getCell(`P${j}`).value = sheetD.getCell(`P3`).value*sheetD.getCell(`N${j}`).value;
                sheetD.getCell(`Q${j}`).value = sheetD.getCell(`Q3`).value*sheetD.getCell(`N${j}`).value;
                sheetD.getCell(`R${j}`).value = sheetD.getCell(`R3`).value*sheetD.getCell(`N${j}`).value;
                sheetD.getCell(`S${j}`).value = sheetD.getCell(`N${j}`).value+sheetD.getCell(`O${j}`).value+sheetD.getCell(`P${j}`).value+sheetD.getCell(`Q${j}`).value+sheetD.getCell(`R${j}`).value;
                sheetD.getCell(`T${j}`).value = sheetD.getCell(`S${j}`).value/total;
                sheetD.mergeCells(`U${j}:V${j}`);
                sheetD.getCell(`U${j}`).value = sheetD.getCell(`V3`).value*sheetD.getCell(`N${j}`).value;
                sheetD.mergeCells(`W${j}:X${j}`);
                sheetD.getCell(`W${j}`).value = sheetD.getCell(`X3`).value*sheetD.getCell(`O${j}`).value;
                sheetD.getCell(`Y${j}`).value = sheetD.getCell(`U${j}`).value+sheetD.getCell(`V${j}`).value;
                sheetD.getCell(`Z${j}`).value = sheetD.getCell(`Y${j}`).value*0.3025;
                sheetD.getCell(`AA${j}`).value = sheetD.getCell(`AA3`).value*sheetD.getCell(`Z${j}`).value;
            }
            else{
                sheetD.getCell(`N${j}`).value = {formula: `H${j}*L${j}`};
                sheetD.getCell(`O${j}`).value = {formula: `O3*N${j}`};
                sheetD.getCell(`P${j}`).value = {formula: `P3*N${j}`};
                sheetD.getCell(`Q${j}`).value = {formula: `Q3*N${j}`};
                sheetD.getCell(`R${j}`).value = {formula: `R3*N${j}`};
                sheetD.getCell(`S${j}`).value = {formula: `N${j}+O${j}+P${j}+Q${j}+R${j}`};
                sheetD.getCell(`T${j}`).value = {formula: `S${j}*2/SUM(S4:S9999)`};
                sheetD.mergeCells(`U${j}:V${j}`);
                sheetD.getCell(`U${j}`).value = {formula: `V3*N${j}`};
                sheetD.mergeCells(`W${j}:X${j}`);
                sheetD.getCell(`W${j}`).value = {formula: `X3*O${j}`};
                sheetD.getCell(`Y${j}`).value = {formula: `U${j}+V${j}`};
                sheetD.getCell(`Z${j}`).value = {formula: `Y${j}*0.3025`};
                sheetD.getCell(`AA${j}`).value = {formula: `AA3*Z${j}`};
            }
        }
    }
    goD();

    // -------------------------------------------------------
    // 5. Sheet E: 基本分析 (這是報告格式，給予通用寬度)
    // -------------------------------------------------------
    sheetE.columns = [
        { header: '項目', key: 'item', width: 15 },
        { header: '說明', key: 'desc', width: 20 },
        { header: '數值', key: 'value', width: 15 },
        { header: '備註', key: 'note', width: 20 },
        // 預留更多欄位給複雜排版
        { key: 'col5', width: 10 },
        { key: 'col6', width: 10 },
        { key: 'col7', width: 10 }
    ];

    // ... 接下來你可以針對各 Sheet 使用 addRow 加資料 ...
    // 例如: sheetA.addRow({ section: '福德', land_no: '123' ... });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "房地產分析報表.xlsx");
    /*


    // 2. 設定欄寬 (Column Width) - 需求 #3
    sheet.columns = [
        { header: '姓名', key: 'name', width: 15 }, // 設定寬度為 15
        { header: '部門', key: 'dept', width: 20 }, // 設定寬度為 20
        { header: '績效', key: 'kpi', width: 10 }
    ];

    // 加入一些資料
    sheet.addRow(['張三', '工程部', 90]);
    sheet.addRow(['李四', '設計部', 85]);
    sheet.addRow(['總計', '', 175]); // 為了示範合併用

    // 3. 設定列高 (Row Height) - 需求 #3
    const firstRow = sheet.getRow(1);
    firstRow.height = 30; // 標題列設高一點

    // 4. 合併儲存格 (Merge Cells) - 需求 #1
    // 把 A4 到 B4 合併 (也就是 '總計' 那一列)
    sheet.mergeCells('A4:B4'); 

    // 5. 設定背景顏色 (Background Color) - 需求 #2
    // 設定標題列 (A1, B1, C1) 的背景色
    ['A1', 'B1', 'C1'].forEach(cellKey => {
        const cell = sheet.getCell(cellKey);
        
        // 設定填滿樣式
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' } // ARGB 格式: FF + Hex顏色 (這裡是淺灰)
        };
        
        // 順便加個粗體和置中比較好看
        cell.font = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // 設定特定格子顏色 (例如 KPI 很高的人給黃色)
    const kpiCell = sheet.getCell('C2'); // 90分那格
    kpiCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' } // 純黃色
    };

    // 6. 輸出檔案
    // ExcelJS 產生的是 Buffer，需要轉成 Blob 讓瀏覽器下載
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "進階報表.xlsx");
    */
}

function mergeCheck(sheet, key, start, end) {
    if (start === end) {
        return;
    }
    let startVal = sheet.getCell(`${key}${start}`).value;
    let startIdx = start;

    for (let i = start + 1; i <= end; i++) {
        let val = sheet.getCell(`${key}${i}`).value;

        if (val !== startVal) {
            if (startVal != null && i - startIdx > 1) {
                sheet.mergeCells(`${key}${startIdx}:${key}${i - 1}`);
            }
            startVal = val;
            startIdx = i;
        } 
        else if (i === end) {
            if (startVal != null && i > startIdx) {
                sheet.mergeCells(`${key}${startIdx}:${key}${i}`);
            }
        }
    }
}

const exportBtn = document.getElementById('export-excel');

if (exportBtn) {
    exportBtn.addEventListener('click', function(event) {
        if (event.shiftKey) {
            exportFancyExcel(true);
            console.log('hi');
        }
        else {
            exportFancyExcel(false);
            console.log('hello');
        }
    });
}