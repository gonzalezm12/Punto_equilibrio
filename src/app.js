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
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            parseExcelData(jsonData);
        } catch (error) {
            console.error("Error reading file:", error);
            alert("Error al procesar el archivo. Asegúrate de que es un Excel o CSV válido.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseExcelData(data) {
    // Asume la primera fila son headers. Buscamos Código y Nombre.
    if (data.length < 2) return;

    // Simplificación: asume que la col 0 es código y col 1 es nombre, o busca por texto
    let headers = data[0].map(h => (h || '').toString().toLowerCase().trim());
    let codeIdx = headers.findIndex(h => h.includes('cod') || h.includes('código'));
    let nameIdx = headers.findIndex(h => h.includes('nom') || h.includes('nombre') || h.includes('desc'));

    if (codeIdx === -1) codeIdx = 0;
    if (nameIdx === -1) nameIdx = 1;

    let newProducts = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const code = row[codeIdx] ? row[codeIdx].toString() : `PROD-${i}`;
        const name = row[nameIdx] ? row[nameIdx].toString() : `Producto ${i}`;

        if (code || name) {
            newProducts.push({
                id: generateId(),
                code: code,
                name: name,
                cv: [],
                cf: [],
                gv: [],
                pvp: 0,
                weight: 0,
                includedInBEP: true
            });
        }
    }

    if (newProducts.length > 0) {
        // Combinar o reemplazar. Aquí optamos por añadir/reemplazar basándonos en código para evitar duplicados.
        newProducts.forEach(newP => {
            const existingIdx = state.products.findIndex(p => p.code === newP.code);
            if (existingIdx >= 0) {
                // Actualiza nombre pero mantiene costos
                state.products[existingIdx].name = newP.name;
            } else {
                state.products.push(newP);
            }
        });
        renderApp();
        alert(`${newProducts.length} productos procesados.`);
    } else {
        alert("No se encontraron datos de productos en el archivo.");
    }
}

// Render Core
function renderApp() {
    renderProductsList();
    renderProductDetails();
    renderGlobalFixedExpenses();
    renderBEPModule();
    renderProjection();
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
            <td>${e.desc}</td>
            <td>${formatMoney(e.amount)}</td>
            <td>
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
    tbody.innerHTML = '';

    let totalFixedProductCosts = 0;
    let totalWeight = 0;
    let totalWeightedMargin = 0;

    state.products.forEach(p => {
        // Calcular costos totales del producto para mostrarlos o usarlos
        const tCV = p.cv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const tCF = p.cf.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const tGV = p.gv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);

        // Sumamos los costos fijos de todos los productos (estén o no incluidos,
        // o solo los incluidos. Asumamos que el PE se calcula con los fijos de los productos incluidos)
        if (p.includedInBEP) {
            totalFixedProductCosts += tCF;
        }

        const margin = p.pvp - tCV - tGV;
        const weightedMargin = margin * p.weight;

        if (p.includedInBEP) {
            totalWeight += p.weight;
            totalWeightedMargin += weightedMargin;
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="checkbox" ${p.includedInBEP ? 'checked' : ''} onchange="app.toggleProductBEP('${p.id}', this.checked)"></td>
            <td>${p.name}</td>
            <td>${formatMoney(tCV)}</td>
            <td><input type="number" value="${p.pvp || ''}" onchange="app.updateProductBEPValue('${p.id}', 'pvp', this.value)" placeholder="PVP" style="width: 80px;" min="0"></td>
            <td><input type="number" value="${(p.weight * 100) || ''}" onchange="app.updateProductBEPValue('${p.id}', 'weight', this.value)" placeholder="%" style="width: 80px;" min="0" max="100"></td>
            <td>${formatMoney(margin)}</td>
            <td>${formatMoney(weightedMargin)}</td>
        `;
        tbody.appendChild(tr);
    });

    // Calcular Total Costos/Gastos Fijos (Globales + Productos incluidos)
    const tGFG = state.globalFixedExpenses.reduce((acc, e) => acc + e.amount, 0);
    const totalFixedCosts = totalFixedProductCosts + tGFG;

    // Calcular Punto de Equilibrio
    let bepUnits = 0;
    if (totalWeightedMargin > 0) {
        bepUnits = totalFixedCosts / totalWeightedMargin;
    }

    // Calcular PE en Dinero (mezcla de ventas)
    let bepMoney = 0;
    state.products.filter(p => p.includedInBEP).forEach(p => {
        const unitsForProduct = bepUnits * p.weight;
        bepMoney += unitsForProduct * p.pvp;
    });

    // Actualizar UI
    document.getElementById('resTotalFixed').textContent = formatMoney(totalFixedCosts);

    const weightElem = document.getElementById('resTotalWeight');
    weightElem.textContent = formatPercent(totalWeight);
    if (Math.abs(totalWeight - 1.0) > 0.001 && totalWeight > 0) {
        weightElem.style.color = '#ff6b6b'; // Red if not exactly 100%
        weightElem.title = 'La ponderación total debería ser 100%';
    } else {
        weightElem.style.color = '';
        weightElem.title = '';
    }

    document.getElementById('resTotalWeightedMargin').textContent = formatMoney(totalWeightedMargin);
    document.getElementById('resBEPUnits').textContent = bepUnits.toFixed(2);
    document.getElementById('resBEPMoney').textContent = formatMoney(bepMoney);
}

// --- Module 4: Projection ---
app.calculateProjection = () => {
    const unitsInput = document.getElementById('projUnits');
    state.projectionUnits = parseFloat(unitsInput.value) || 0;
    renderProjection();
};

function renderProjection() {
    const units = state.projectionUnits;
    const includedProducts = state.products.filter(p => p.includedInBEP);

    let totalRevenue = 0;
    let totalCV = 0;
    let totalGV = 0;
    let totalCF = 0; // Costos fijos asignados
    let totalAcquisitionCosts = 0;

    const tGFG = state.globalFixedExpenses.reduce((acc, e) => acc + e.amount, 0);

    const tbody = document.querySelector('#projProductTable tbody');
    tbody.innerHTML = '';

    includedProducts.forEach(p => {
        // Unidades proyectadas para este producto basado en su ponderación
        const productUnits = units * p.weight;

        // Suma de costos/gastos unitarios
        const unitCV = p.cv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const unitCF = p.cf.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
        const unitGV = p.gv.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);

        const revenue = productUnits * p.pvp;
        const cv = productUnits * unitCV;
        const gv = productUnits * unitGV;

        // Asumimos que los CF del producto ya son totales y no por unidad,
        // O si son por unidad, multiplicamos. Normalmente los CF son totales por producto.
        // Asumiremos que el input del usuario es el CF total para la producción.
        const cf = unitCF; // Si fueran unitarios: productUnits * unitCF;

        // Costo de adquisición unitario = (CV + CF + GV) / (si se requiere) o simplemente mostrar la suma de unitarios
        // Si CF es total, costo unitario CF = CF / productUnits
        const unitCFAllocated = productUnits > 0 ? (cf / productUnits) : 0;
        const unitAcquisitionCost = unitCV + unitCFAllocated + unitGV;

        totalRevenue += revenue;
        totalCV += cv;
        totalGV += gv;
        totalCF += cf;

        // Investment (Total costs)
        const productInvestment = cv + cf + gv;
        totalAcquisitionCosts += productInvestment;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${p.name}</td>
            <td>${productUnits.toFixed(2)}</td>
            <td>${formatMoney(revenue)}</td>
            <td>${formatMoney(unitAcquisitionCost)}</td>
        `;
        tbody.appendChild(tr);
    });

    const margin = totalRevenue - totalCV - totalGV;
    const investment = totalCV + totalGV + totalCF + tGFG;
    const profit = margin - totalCF - tGFG;
    const roi = investment > 0 ? (profit / investment) * 100 : 0;

    document.getElementById('projRevenue').textContent = formatMoney(totalRevenue);
    document.getElementById('projCV').textContent = formatMoney(totalCV);
    document.getElementById('projGV').textContent = formatMoney(totalGV);
    document.getElementById('projMargin').textContent = formatMoney(margin);
    document.getElementById('projCFA').textContent = formatMoney(totalCF);
    document.getElementById('projGFG').textContent = formatMoney(tGFG);
    document.getElementById('projInvestment').textContent = formatMoney(investment);

    const profitElem = document.getElementById('projProfit');
    profitElem.textContent = formatMoney(profit);
    profitElem.style.color = profit >= 0 ? 'var(--accent-hover)' : '#ff6b6b';

    const roiElem = document.getElementById('projROI');
    roiElem.textContent = roi.toFixed(2) + '%';
    roiElem.style.color = roi >= 0 ? 'var(--accent-hover)' : '#ff6b6b';
}

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
            <td>${sanitizeHTML(p.code)}</td>
            <td>${sanitizeHTML(p.name)}</td>
            <td>
                <button class="btn-sm" onclick="app.selectProduct('${p.id}')">Detalles</button>
                <button class="btn-danger" onclick="app.deleteProduct('${p.id}')">Eliminar</button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const summaryPanel = document.getElementById('portfolioSummary');
    if (state.products.length > 0) {
        summaryPanel.classList.remove('hidden');
        document.getElementById('portTotalCV').textContent = portCV.toFixed(2);
        document.getElementById('portTotalCF').textContent = portCF.toFixed(2);
        document.getElementById('portTotalGV').textContent = portGV.toFixed(2);
    } else {
        summaryPanel.classList.add('hidden');
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
    renderBEPModule();
    renderProjection();
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
    alert("Detalles guardados. (El guardado es automático en memoria)");
};
