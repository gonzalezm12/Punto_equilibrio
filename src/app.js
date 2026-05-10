// Central State Management
const state = {
    products: [], // { id, code, name, cv: [], cf: [], gv: [], pvp: 0, weight: 0, includedInBEP: true }
    globalFixedExpenses: [], // { id, desc, amount }
    selectedProductId: null,
    projectionUnits: 0
};

// Utils
const generateId = () => '_' + Math.random().toString(36).substr(2, 9);
const formatMoney = (amount) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(amount);
const formatPercent = (val) => (val * 100).toFixed(2) + '%';

// DOM Elements Initialization
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initDropzone();
    renderApp();
});

// --- Tab Logic ---
function initTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('data-tab');

            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));

            btn.classList.add('active');
            document.getElementById(targetId).classList.add('active');
        });
    });
}

window.app = {};

// --- Module 1: Excel Upload & Product Management ---
function initDropzone() {
    const dropzone = document.getElementById('excelDropzone');
    const fileInput = document.getElementById('fileInput');

    dropzone.addEventListener('click', () => fileInput.click());

    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFileUpload(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFileUpload(e.target.files[0]);
        }
    });
}

function handleFileUpload(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Leer Hoja 1: Productos
            const sheet1Name = workbook.SheetNames[0];
            const productsData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet1Name], { header: 1 });
            
            // Leer Hoja 2: Detalle_Costos (si existe)
            let detailsData = null;
            if (workbook.SheetNames.length > 1) {
                const sheet2Name = workbook.SheetNames.find(n => n.toLowerCase().includes('detalle') || n.toLowerCase().includes('costo'));
                if (sheet2Name) {
                    detailsData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet2Name], { header: 1 });
                }
            }

            parseExcelData(productsData, detailsData);
        } catch (error) {
            console.error("Error reading file:", error);
            alert("Error al procesar el archivo. Asegúrate de usar la plantilla oficial.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseExcelData(productsData, detailsData) {
    if (!productsData || productsData.length < 2) return;

    // Función para limpiar texto (quitar tildes, espacios y minúsculas)
    const clean = (str) => (str || '').toString().toLowerCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

    // --- 1. Procesar Productos (Hoja 1) ---
    let headerRowIdx = -1;
    for (let i = 0; i < productsData.length; i++) {
        const row = productsData[i].map(c => clean(c));
        if (row.some(c => c.includes('cod')) || row.some(c => c.includes('nom'))) {
            headerRowIdx = i;
            break;
        }
    }

    if (headerRowIdx === -1) {
        alert("No se encontró la fila de cabeceras en la hoja de Productos.");
        return;
    }

    const headers = productsData[headerRowIdx].map(h => clean(h));
    const codeIdx = headers.findIndex(h => h.includes('cod'));
    const nameIdx = headers.findIndex(h => h.includes('nom'));
    const pvpIdx = headers.findIndex(h => h.includes('precio') || h.includes('pvp'));

    let addedCount = 0;
    let updatedCount = 0;

    for (let i = headerRowIdx + 1; i < productsData.length; i++) {
        const row = productsData[i];
        if (!row || row.length === 0) continue;

        const rawCode = codeIdx !== -1 && row[codeIdx] ? row[codeIdx].toString().trim() : '';
        const nameStr = nameIdx !== -1 && row[nameIdx] ? row[nameIdx].toString().trim() : '';
        const pvp = pvpIdx !== -1 ? parseFloat(row[pvpIdx]) || 0 : 0;

        if (!rawCode && !nameStr) continue;

        const code = rawCode.toUpperCase() || `PROD-${generateId().substring(0,4).toUpperCase()}`;
        const name = nameStr || `Producto Sin Nombre`;

        let targetProduct = state.products.find(p => (p.code || '').toUpperCase() === code);
        
        if (targetProduct) {
            targetProduct.name = name;
            targetProduct.cv = []; targetProduct.cf = []; targetProduct.gv = [];
            updatedCount++;
        } else {
            targetProduct = {
                id: generateId(),
                code: code,
                name: name,
                cv: [], cf: [], gv: [],
                pvp: pvp,
                weight: 0,
                projectedUnits: 0,
                includedInBEP: true
            };
            state.products.push(targetProduct);
            addedCount++;
        }
        if (pvp > 0) targetProduct.pvp = pvp;
    }

    // --- 2. Procesar Detalles (Hoja 2) ---
    if (detailsData && detailsData.length > 1) {
        let detHeaderIdx = -1;
        for (let i = 0; i < detailsData.length; i++) {
            const row = detailsData[i].map(c => clean(c));
            if (row.some(c => c.includes('cod')) || row.some(c => c.includes('tipo'))) {
                detHeaderIdx = i;
                break;
            }
        }

        if (detHeaderIdx !== -1) {
            const detHeaders = detailsData[detHeaderIdx].map(h => clean(h));
            const dCodeIdx = detHeaders.findIndex(h => h.includes('cod'));
            const dTypeIdx = detHeaders.findIndex(h => h.includes('tipo'));
            const dDescIdx = detHeaders.findIndex(h => h.includes('desc'));
            const dAmountIdx = detHeaders.findIndex(h => h.includes('monto') || h.includes('valor'));

            for (let i = detHeaderIdx + 1; i < detailsData.length; i++) {
                const row = detailsData[i];
                if (!row || row.length < 2) continue;

                const rawCode = (row[dCodeIdx] || '').toString().trim().toUpperCase();
                const type = clean(row[dTypeIdx]);
                const desc = (row[dDescIdx] || 'Importado').toString().trim();
                const amount = parseFloat(row[dAmountIdx]) || 0;

                const product = state.products.find(p => p.code.toUpperCase() === rawCode);
                if (product && amount > 0) {
                    if (type === 'cv' || (type.includes('costo') && type.includes('variable'))) {
                        product.cv.push({ desc, amount });
                    } else if (type === 'cf' || type.includes('fijo')) {
                        product.cf.push({ desc, amount });
                    } else if (type === 'gv' || (type.includes('gasto') && type.includes('variable'))) {
                        product.gv.push({ desc, amount });
                    }
                }
            }
        }
    }

    renderApp();
    alert(`Importación completada: ${addedCount} nuevos, ${updatedCount} actualizados.`);
}

app.downloadTemplate = () => {
    // Hoja 1: Productos
    const ws_products_data = [
        ["Código", "Nombre", "Precio de Venta"],
        ["PR001", "Producto de Ejemplo A", 100.00],
        ["PR002", "Producto de Ejemplo B", 200.00]
    ];
    const ws_products = XLSX.utils.aoa_to_sheet(ws_products_data);

    // Hoja 2: Detalle_Costos
    const ws_details_data = [
        ["Código Producto", "Tipo (CV / CF / GV)", "Descripción", "Monto"],
        ["PR001", "CV", "Tela de Algodón", 25.00],
        ["PR001", "CV", "Hilos y Botones", 5.50],
        ["PR001", "CF", "Etiqueta Marca", 1.20],
        ["PR002", "CV", "Cuero Sintético", 50.00],
        ["PR002", "GV", "Comisión de Marketplace", 15.00]
    ];
    const ws_details = XLSX.utils.aoa_to_sheet(ws_details_data);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws_products, "Productos");
    XLSX.utils.book_append_sheet(wb, ws_details, "Detalle_Costos");

    XLSX.writeFile(wb, "Plantilla_Avanzada_Punto_Equilibrio.xlsx");
};

app.addManualProduct = () => {
    const codeInput = document.getElementById('manualCode');
    const nameInput = document.getElementById('manualName');
    const pvpInput = document.getElementById('manualPVP');
    const code = codeInput.value.trim();
    const name = nameInput.value.trim();
    const pvp = parseFloat(pvpInput.value) || 0;

    if (!code || !name || pvp <= 0) {
        alert("Por favor, ingrese un Código, un Nombre y un Precio de Venta (PVP) mayor a 0.");
        return;
    }

    if (state.products.find(p => p.code === code)) {
        alert("Ya existe un producto con este código.");
        return;
    }

    state.products.push({
        id: generateId(),
        code: code,
        name: name,
        cv: [],
        cf: [],
        gv: [],
        pvp: pvp,
        weight: 0,
        projectedUnits: 0,
        includedInBEP: true
    });

    codeInput.value = '';
    nameInput.value = '';
    pvpInput.value = '';
    renderApp();
};

// Render Core
function renderApp() {
    renderProductsList();
    renderProductDetails();
    renderGlobalFixedExpenses();
    renderBEPModule();
    app.renderProjection();
}

// --- Module 2: Global Fixed Expenses ---
app.addGlobalFixedExpense = () => {
    const descInput = document.getElementById('gfgDesc');
    const amountInput = document.getElementById('gfgAmount');
    const desc = descInput.value.trim();
    const amount = parseFloat(amountInput.value) || 0;

    if (desc && amount > 0) {
        state.globalFixedExpenses.push({
            id: generateId(),
            desc: desc,
            amount: amount
        });
        descInput.value = '';
        amountInput.value = '';
        renderApp();
    } else {
        alert("Por favor ingrese una descripción y un monto mayor a 0.");
    }
};

app.removeGlobalFixedExpense = (id) => {
    state.globalFixedExpenses = state.globalFixedExpenses.filter(e => e.id !== id);
    renderApp();
};

function renderGlobalFixedExpenses() {
    const tbody = document.querySelector('#gfgTable tbody');
    tbody.innerHTML = '';

    let total = 0;
    state.globalFixedExpenses.forEach(e => {
        total += e.amount;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td data-label="Descripción">${e.desc}</td>
            <td data-label="Monto">${formatMoney(e.amount)}</td>
            <td data-label="Acción">
                <button class="btn-danger" onclick="app.removeGlobalFixedExpense('${e.id}')">Eliminar</button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('totalGFG').textContent = formatMoney(total);
}

// --- Module 3: Weighted Break-Even Point ---
app.toggleProductBEP = (id, isChecked) => {
    const product = state.products.find(p => p.id === id);
    if (product) {
        product.includedInBEP = isChecked;
        renderApp();
    }
};

app.updateProductBEPValue = (id, field, value) => {
    const product = state.products.find(p => p.id === id);
    if (product) {
        let numValue = parseFloat(value) || 0;
        if (field === 'weight') {
            numValue = numValue / 100; // Convertir de porcentaje a decimal (ej 50 -> 0.5)
        }
        product[field] = numValue;
        renderApp();
    }
};

function renderBEPModule() {
    const tbody = document.querySelector('#pepTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    let totalFixedProductCosts = 0;
    let totalWeight = 0;
    let totalWeightedMargin = 0;

    // Calcular Total Costos/Gastos Fijos (Globales + Productos incluidos)
    const tGFG = state.globalFixedExpenses.reduce((acc, e) => acc + e.amount, 0);

    // Primera pasada: Calcular costos fijos totales y margen ponderado total
    state.products.forEach(p => {
        const tCF = p.cf.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        if (p.includedInBEP) {
            totalFixedProductCosts += tCF;
            const tCV = p.cv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
            const tGV = p.gv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
            const margin = p.pvp - tCV - tGV;
            const weightedMargin = margin * p.weight;
            totalWeight += p.weight;
            totalWeightedMargin += weightedMargin;
        }
    });

    const totalFixedCosts = totalFixedProductCosts + tGFG;

    // Calcular Punto de Equilibrio Global en Unidades
    let globalBepUnits = 0;
    if (totalWeightedMargin > 0) {
        globalBepUnits = totalFixedCosts / totalWeightedMargin;
    }

    let globalBepMoney = 0;

    // Segunda pasada: Renderizar tabla y calcular BEP por producto
    state.products.forEach(p => {
        const tCV = p.cv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const tGV = p.gv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const totalVariables = tCV + tGV;
        const margin = p.pvp - totalVariables;
        const weightedMarginPerProduct = margin * p.weight;

        let bepUnidades = 0;
        let bepDinero = 0;

        if (p.includedInBEP && margin > 0) {
            bepUnidades = globalBepUnits * p.weight;
            bepDinero = bepUnidades * p.pvp;
            globalBepMoney += bepDinero;
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td data-label="Incluir"><input type="checkbox" ${p.includedInBEP ? 'checked' : ''} onchange="app.toggleProductBEP('${p.id}', this.checked)"></td>
            <td data-label="Producto">${p.name}</td>
            <td data-label="Costo Var.">${formatMoney(tCV)}</td>
            <td data-label="Gasto Var.">${formatMoney(tGV)}</td>
            <td data-label="Total Var.">${formatMoney(totalVariables)}</td>
            <td data-label="PVP ($)"><input type="number" value="${p.pvp || ''}" onchange="app.updateProductBEPValue('${p.id}', 'pvp', this.value)" placeholder="PVP" style="width: 80px;" min="0"></td>
            <td data-label="Margen Contrib." class="font-bold ${margin < 0 ? 'text-red-500' : ''}">${formatMoney(margin)}</td>
            <td data-label="Ponderación (%)"><input type="number" value="${(p.weight * 100) || ''}" onchange="app.updateProductBEPValue('${p.id}', 'weight', this.value)" placeholder="%" style="width: 80px;" min="0" max="100"></td>
            <td data-label="M. Contrib. Pond." class="font-bold text-blue-500">${formatMoney(weightedMarginPerProduct)}</td>
            <td data-label="PE (Unidades)">${bepUnidades.toFixed(2)}</td>
            <td data-label="PE ($)">${formatMoney(bepDinero)}</td>
        `;
        tbody.appendChild(tr);
    });

    // Actualizar Resúmenes en la UI
    document.getElementById('resTotalFixed').textContent = formatMoney(totalFixedCosts);
    const weightElem = document.getElementById('resTotalWeight');
    weightElem.textContent = formatPercent(totalWeight);
    
    if (Math.abs(totalWeight - 1.0) > 0.001 && totalWeight > 0) {
        weightElem.style.color = '#ff6b6b';
    } else {
        weightElem.style.color = '';
    }

    document.getElementById('resTotalWeightedMargin').textContent = formatMoney(totalWeightedMargin || 0);
    document.getElementById('resBEPUnits').textContent = (globalBepUnits || 0).toFixed(2);
    document.getElementById('resBEPMoney').textContent = formatMoney(globalBepMoney || 0);
}


// --- Module 4: Projection ---
app.updateProductProjection = (id, value) => {
    const product = state.products.find(p => p.id === id);
    if (product) {
        product.projectedUnits = parseFloat(value) || 0;
        app.renderProjection();
    }
};

app.renderProjection = function() {
    const inputBody = document.getElementById('projInputBody');
    const resultBody = document.getElementById('projResultBody');
    if (!inputBody || !resultBody) return;

    inputBody.innerHTML = '';
    resultBody.innerHTML = '';

    const includedProducts = state.products.filter(p => p.includedInBEP);
    const globalFixed = state.globalFixedExpenses.reduce((acc, e) => acc + (parseFloat(e.amount) || 0), 0);

    let totalRevenue = 0;
    let totalCV = 0;
    let totalGV = 0;
    let totalCF = 0;
    let totalInvestment = 0;
    let totalProfit = 0;

    // Calcular BEP Global para referencias
    let globalBepUnits = 0;
    let totalWeightedMargin = 0;
    let totalFixedCostsBase = globalFixed;

    includedProducts.forEach(p => {
        const tCF = p.cf.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const tCV = p.cv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const tGV = p.gv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        totalFixedCostsBase += tCF;
        totalWeightedMargin += (p.pvp - tCV - tGV) * (p.weight || 0);
    });

    if (totalWeightedMargin > 0) {
        globalBepUnits = totalFixedCostsBase / totalWeightedMargin;
    }

    includedProducts.forEach(p => {
        const productUnits = p.projectedUnits || 0;
        const unitCV = p.cv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const unitGV = p.gv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const unitCF = p.cf.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        
        // Cálculos Individuales con Alícuota basada en PE
        const unitMargin = p.pvp - unitCV - unitGV;
        const bepUnitsPerProduct = globalBepUnits * (p.weight || 0);
        
        // La Alícuota es lo que el producto debe cubrir para llegar al PE
        // (PE del producto * Margen de Contribución)
        const targetFixedCoverage = bepUnitsPerProduct * unitMargin;
        
        const revenue = productUnits * p.pvp;
        const cv = productUnits * unitCV;
        const gv = productUnits * unitGV;
        const cf = unitCF; // Costo fijo específico del producto
        const aliquotGF = targetFixedCoverage - unitCF; // Diferencia para cubrir la cuota de fijos globales
        
        const investment = cv + gv + cf + aliquotGF;
        const profitability = revenue - investment;
        const roi = investment > 0 ? (profitability / investment) * 100 : 0;
        const margin = revenue > 0 ? (profitability / revenue) * 100 : 0;
        const acquisitionCost = productUnits > 0 ? (investment / productUnits) : (unitCV + unitGV + (targetFixedCoverage / (bepUnitsPerProduct || 1)));

        // Acumuladores Globales
        totalRevenue += revenue;
        totalCV += cv;
        totalGV += gv;
        totalCF += cf;
        totalInvestment += investment;
        totalProfit += profitability;

        // Renderizar fila de ENTRADA
        const bepUnidades = globalBepUnits * p.weight;
        let statusHtml = '<span class="status-badge status-none">Sin Meta</span>';
        if (productUnits > 0) {
            if (productUnits >= bepUnidades) {
                statusHtml = '<span class="status-badge status-success">Meta Alcanzada</span>';
            } else {
                statusHtml = '<span class="status-badge status-danger">Bajo Equilibrio</span>';
            }
        }

        const inputTr = document.createElement('tr');
        inputTr.innerHTML = `
            <td data-label="Producto">${p.name}</td>
            <td data-label="PVP ($)">${formatMoney(p.pvp || 0)}</td>
            <td data-label="PE Referencia" class="font-bold text-blue-500">${(bepUnidades || 0).toFixed(2)}</td>
            <td data-label="Unidades Proyectadas">
                <input type="number" value="${p.projectedUnits || ''}" 
                       onchange="app.updateProductProjection('${p.id}', this.value)" 
                       placeholder="0" class="input-small" style="width: 80px;" min="0">
            </td>
            <td data-label="Estado">${statusHtml}</td>
        `;
        inputBody.appendChild(inputTr);

        // Renderizar fila de RESULTADO
        const resultTr = document.createElement('tr');
        resultTr.innerHTML = `
            <td data-label="Producto">${p.name || 'Sin Nombre'}</td>
            <td data-label="Ingresos">${formatMoney(revenue)}</td>
            <td data-label="Costos Var.">${formatMoney(cv)}</td>
            <td data-label="Gastos Var.">${formatMoney(gv)}</td>
            <td data-label="Costos Fijos">${formatMoney(cf)}</td>
            <td data-label="Alícuota G.F.">${formatMoney(aliquotGF)}</td>
            <td data-label="Inversión">${formatMoney(investment)}</td>
            <td data-label="Rentabilidad" class="${profitability >= 0 ? 'text-green' : 'text-red'} font-bold">${formatMoney(profitability)}</td>
            <td data-label="ROI (%)">${roi.toFixed(2)}%</td>
            <td data-label="Margen (%)">${margin.toFixed(2)}%</td>
            <td data-label="Costo Adq.">${formatMoney(acquisitionCost)}</td>
        `;
        resultBody.appendChild(resultTr);
    });

    // Actualizar Resumen Global
    const totalROI = totalInvestment > 0 ? (totalProfit / totalInvestment) * 100 : 0;
    const totalMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;

    document.getElementById('totalFixedProj').textContent = formatMoney(totalCF + globalFixed);
    document.getElementById('totalRevenueProj').textContent = formatMoney(totalRevenue);
    document.getElementById('totalInvestmentProj').textContent = formatMoney(totalInvestment);
    document.getElementById('totalProfitProj').textContent = formatMoney(totalProfit);
    document.getElementById('totalROIProj').textContent = totalROI.toFixed(2) + '%';
    document.getElementById('totalMarginProj').textContent = totalMargin.toFixed(2) + '%';
};

app.generatePDF = function() {
    if (!window.jspdf) {
        alert("Las librerías de PDF no se han cargado correctamente. Por favor, recarga la página.");
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4'); // Vertical A4 para reporte formal
    const pageWidth = doc.internal.pageSize.getWidth();
    const primaryColor = [60, 179, 113];

    // --- CABECERA ---
    doc.setFontSize(22);
    doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
    doc.text("DETRÁS DEL BALANCE", 14, 20);
    
    doc.setFontSize(14);
    doc.setTextColor(100, 100, 100);
    doc.text("Reporte Financiero Integral", 14, 28);

    doc.setFontSize(9);
    doc.setTextColor(0, 0, 0);
    doc.text("Instagram: @detras_del_balance", 14, 38);
    doc.text("WhatsApp: +58 412-3422609", 14, 43);
    doc.text(`Fecha de Emisión: ${new Date().toLocaleDateString()}`, pageWidth - 14, 38, { align: 'right' });

    let currentY = 55;

    // --- SECCIÓN 1: PRODUCTOS Y COSTOS ---
    doc.setFontSize(12);
    doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
    doc.text("1. ESTRUCTURA DE COSTOS POR PRODUCTO", 14, currentY);

    const productData = state.products.map(p => {
        const sum = (arr) => (arr || []).reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const tCV = sum(p.cv);
        const tCF = sum(p.cf);
        const tGV = sum(p.gv);
        return [p.name || 'S/N', formatMoney(p.pvp), formatMoney(tCV), formatMoney(tGV), formatMoney(tCF)];
    });

    doc.autoTable({
        startY: currentY + 5,
        head: [['Producto', 'PVP', 'Costo Var.', 'Gasto Var.', 'Costo Fijo Asig.']],
        body: productData,
        theme: 'striped',
        headStyles: { fillColor: primaryColor }
    });

    currentY = doc.lastAutoTable.finalY + 15;

    // --- SECCIÓN 2: GASTOS FIJOS GLOBALES ---
    doc.text("2. GASTOS FIJOS GLOBALES", 14, currentY);
    
    const gfData = state.globalFixedExpenses.map(e => [e.desc, formatMoney(e.amount)]);
    const totalGF = state.globalFixedExpenses.reduce((sum, e) => sum + e.amount, 0);
    gfData.push([{ content: 'TOTAL GASTOS FIJOS GLOBALES', styles: { fontStyle: 'bold' } }, { content: formatMoney(totalGF), styles: { fontStyle: 'bold' } }]);

    doc.autoTable({
        startY: currentY + 5,
        head: [['Descripción', 'Monto']],
        body: gfData,
        theme: 'plain',
        headStyles: { fillColor: [100, 100, 100] }
    });

    currentY = doc.lastAutoTable.finalY + 15;

    // --- SECCIÓN 3: PUNTO DE EQUILIBRIO PONDERADO ---
    if (currentY > 230) { doc.addPage(); currentY = 20; }
    doc.text("3. ANÁLISIS DE PUNTO DE EQUILIBRIO PONDERADO", 14, currentY);

    const bepBody = [];
    const bepRows = document.querySelectorAll('#pepTable tbody tr');
    bepRows.forEach(row => {
        const cells = Array.from(row.querySelectorAll('td')).map(td => {
            // Si el td contiene un input, tomar su valor, si no, tomar el texto
            const input = td.querySelector('input');
            return input ? input.value : td.textContent.trim();
        });
        // Filtrar solo las columnas que queremos en el PDF (Producto, Ponderación, Margen, PE Unidades, PE Ventas)
        // Estructura tabla: [Incluir, Producto, CV, GV, TotalVar, PVP, Margen, Pond, MargenPond, PEUnits, PEVentas]
        const filteredCells = [
            cells[1], // Producto
            cells[7] + '%', // Ponderación
            cells[6], // Margen Contrib.
            cells[9], // PE Unidades
            cells[10] // PE Ventas
        ];
        bepBody.push(filteredCells);
    });

    doc.autoTable({
        startY: currentY + 5,
        head: [['Producto', 'Ponderación', 'Margen Contrib.', 'PE (Unidades)', 'PE (Ventas $)']],
        body: bepBody,
        theme: 'grid',
        headStyles: { fillColor: primaryColor }
    });

    // Resumen Global de PE
    const resBEPMoney = document.getElementById('resBEPMoney');
    const resTotalWeightedMargin = document.getElementById('resTotalWeightedMargin');
    
    const globalBESummary = [
        ["Ventas Totales en Equilibrio", resBEPMoney ? resBEPMoney.textContent : "$0.00"],
        ["Margen de Contribución Ponderado", resTotalWeightedMargin ? resTotalWeightedMargin.textContent : "$0.00"]
    ];

    doc.autoTable({
        startY: doc.lastAutoTable.finalY + 5,
        body: globalBESummary,
        theme: 'plain',
        styles: { fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 100 } }
    });

    doc.addPage();
    currentY = 20;

    // --- SECCIÓN 4: SIMULACIÓN DE PROYECCIÓN ---
    doc.text("4. SIMULACIÓN DE PROYECCIÓN Y RENTABILIDAD", 14, currentY);

    const projBody = [];
    document.querySelectorAll('#projResultBody tr').forEach(row => {
        projBody.push(Array.from(row.querySelectorAll('td')).map(td => td.textContent));
    });

    doc.autoTable({
        startY: currentY + 5,
        head: [['Producto', 'Ingresos', 'Costos Var.', 'Gastos Var.', 'Costos Fijos', 'Alicuota G.F.', 'Inversión', 'Rentabilidad', 'ROI (%)', 'Margen (%)', 'Costo Adq.']],
        body: projBody,
        theme: 'grid',
        headStyles: { fillColor: primaryColor, fontSize: 7 },
        styles: { fontSize: 7 }
    });

    // --- RESUMEN FINAL ---
    doc.setFontSize(14);
    doc.text("CONCLUSIONES GENERALES", 14, doc.lastAutoTable.finalY + 15);

    const finalSummary = [
        ["Ingresos Brutos Proyectados", document.getElementById('totalRevenueProj').textContent],
        ["Inversión Total Operativa", document.getElementById('totalInvestmentProj').textContent],
        ["UTILIDAD NETA ESTIMADA", document.getElementById('totalProfitProj').textContent],
        ["RETORNO DE INVERSIÓN (ROI)", document.getElementById('totalROIProj').textContent],
        ["MARGEN DE UTILIDAD SOBRE VENTAS", document.getElementById('totalMarginProj').textContent]
    ];

    doc.autoTable({
        startY: doc.lastAutoTable.finalY + 20,
        body: finalSummary,
        theme: 'striped',
        styles: { fontSize: 10, cellPadding: 5 },
        columnStyles: { 0: { fontStyle: 'bold', cellWidth: 100 } }
    });

    // Pie de página
    doc.setFontSize(8);
    doc.setTextColor(150, 150, 150);
    doc.text("Este documento es un análisis financiero proyectado. Los resultados dependen del cumplimiento de las metas de ventas.", pageWidth / 2, doc.internal.pageSize.getHeight() - 10, { align: 'center' });

    doc.save(`Informe_Financiero_DetrasDelBalance_${new Date().toISOString().slice(0,10)}.pdf`);
};

// --- Utils ---
const sanitizeHTML = (str) => {
    const temp = document.createElement('div');
    temp.textContent = str;
    return temp.innerHTML;
};

// --- Module 1 Rendering ---
function renderProductsList() {
    const tbody = document.querySelector('#productsTable tbody');
    tbody.innerHTML = '';

    let portCV = 0;
    let portCF = 0;
    let portGV = 0;

    state.products.forEach(p => {
        const sum = (arr) => arr.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        portCV += sum(p.cv);
        portCF += sum(p.cf);
        portGV += sum(p.gv);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td data-label="Código">${sanitizeHTML(p.code)}</td>
            <td data-label="Nombre">${sanitizeHTML(p.name)}</td>
            <td data-label="Acción">
                <button class="btn-sm" onclick="app.selectProduct('${p.id}')">Detalles</button>
                <button class="btn-danger" onclick="app.deleteProduct('${p.id}')">Eliminar</button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const summaryPanel = document.getElementById('portfolioSummary');
    const emptyState = document.getElementById('emptyState');
    const productsTable = document.getElementById('productsTable');

    if (state.products.length > 0) {
        summaryPanel.classList.remove('hidden');
        emptyState.classList.add('hidden');
        productsTable.classList.remove('hidden');
        document.getElementById('portTotalCV').textContent = portCV.toFixed(2);
        document.getElementById('portTotalCF').textContent = portCF.toFixed(2);
        document.getElementById('portTotalGV').textContent = portGV.toFixed(2);
    } else {
        summaryPanel.classList.add('hidden');
        emptyState.classList.remove('hidden');
        productsTable.classList.add('hidden');
    }
}

app.selectProduct = (id) => {
    state.selectedProductId = id;
    renderApp();
};

app.deleteProduct = (id) => {
    if(confirm("¿Seguro que deseas eliminar este producto?")) {
        state.products = state.products.filter(p => p.id !== id);
        if (state.selectedProductId === id) state.selectedProductId = null;
        renderApp();
    }
};

function renderProductDetails() {
    const container = document.getElementById('productDetailsForm');
    const noProductMsg = document.getElementById('noProductSelected');
    const nameSpan = document.getElementById('selectedProductName');

    if (!state.selectedProductId) {
        container.classList.add('hidden');
        noProductMsg.classList.remove('hidden');
        nameSpan.textContent = '';
        return;
    }

    const product = state.products.find(p => p.id === state.selectedProductId);
    if (!product) return;

    container.classList.remove('hidden');
    noProductMsg.classList.add('hidden');
    nameSpan.textContent = `- ${product.name} (${product.code})`;

    renderCostSection('cvContainer', product.cv);
    renderCostSection('cfContainer', product.cf);
    renderCostSection('gvContainer', product.gv);

    updateProductTotals(product);
}

function renderCostSection(containerId, items) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';
    items.forEach((item, index) => {
        const div = document.createElement('div');
        div.className = 'cost-row';
        div.innerHTML = `
            <input type="text" value="${item.desc}" onchange="app.updateCostItem('${containerId}', ${index}, 'desc', this.value)" placeholder="Descripción">
            <input type="number" value="${item.amount}" onchange="app.updateCostItem('${containerId}', ${index}, 'amount', this.value)" placeholder="Monto" min="0" step="0.01">
            <button class="btn-danger" onclick="app.removeCostRow('${containerId}', ${index})">X</button>
        `;
        container.appendChild(div);
    });
}

app.addCostRow = (containerId) => {
    if (!state.selectedProductId) return;
    const product = state.products.find(p => p.id === state.selectedProductId);
    const targetArray = getTargetArray(product, containerId);

    targetArray.push({ desc: '', amount: 0 });
    renderProductDetails();
};

app.removeCostRow = (containerId, index) => {
    const product = state.products.find(p => p.id === state.selectedProductId);
    const targetArray = getTargetArray(product, containerId);
    targetArray.splice(index, 1);
    renderProductDetails();
};

app.updateCostItem = (containerId, index, field, value) => {
    const product = state.products.find(p => p.id === state.selectedProductId);
    const targetArray = getTargetArray(product, containerId);
    if (field === 'amount') {
        targetArray[index][field] = parseFloat(value) || 0;
    } else {
        targetArray[index][field] = value;
    }
    updateProductTotals(product);
    renderProductsList(); // Update portfolio summary
    renderBEPModule();
    app.renderProjection();
};

function getTargetArray(product, containerId) {
    if (containerId === 'cvContainer') return product.cv;
    if (containerId === 'cfContainer') return product.cf;
    if (containerId === 'gvContainer') return product.gv;
    return [];
}

function updateProductTotals(product) {
    const sum = (arr) => arr.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
    const tCV = sum(product.cv);
    const tCF = sum(product.cf);
    const tGV = sum(product.gv);

    document.getElementById('totalCV').textContent = tCV.toFixed(2);
    document.getElementById('totalCF').textContent = tCF.toFixed(2);
    document.getElementById('totalGV').textContent = tGV.toFixed(2);
}

app.saveProductDetails = () => {
    renderApp();
    alert("Detalles guardados y resumen actualizado.");
};
