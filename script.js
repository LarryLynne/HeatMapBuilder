// --- 1. ИНИЦИАЛИЗАЦИЯ КАРТЫ И СЛОЕВ ---
const lightMap = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { attribution: '&copy; OSM', maxZoom: 19, crossOrigin: true });
const darkMap = L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', { attribution: '&copy; CartoDB', maxZoom: 20, crossOrigin: true });
const satelliteMap = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { attribution: '&copy; Esri', maxZoom: 18, crossOrigin: true });

const map = L.map('map', { center: [49.589, 34.551], zoom: 6, layers: [lightMap], preferCanvas: true });
L.control.layers({ "Світла": lightMap, "Темна": darkMap, "Супутник": satelliteMap }, null, {position: 'topright'}).addTo(map);

// Инициализация инструментов рисования Geoman
map.pm.addControls({
    position: 'topleft', drawPolygon: true, drawRectangle: true, drawMarker: false,
    drawCircleMarker: false, drawPolyline: false, drawCircle: false, drawText: false,
    editMode: false, dragMode: false, cutPolygon: false, removalMode: true
});
map.pm.setLang('ua');

const screenshoter = L.simpleMapScreenshoter({ hidden: true, mimeType: 'image/png' }).addTo(map);

// --- 2. ПЕРЕМЕННЫЕ СОСТОЯНИЯ ---
let heatLayer = null;
let cachedData = [];
let activeFilters = {};
let currentDrawnLayer = null;

const radiusInput = document.getElementById('radiusRange');
const blurInput = document.getElementById('blurRange');
const maxInput = document.getElementById('maxRange');
const gradientSelect = document.getElementById('gradientSelect');
const statusDiv = document.getElementById('status');
const uploadLabel = document.getElementById('uploadLabel');
const snapBtn = document.getElementById('snapBtn');
const dynamicSlicers = document.getElementById('dynamicSlicers');

// Элементы окон
const zoneModal = document.getElementById('zoneModal');
const zoneTotalVal = document.getElementById('zoneTotalVal');
const zoneListContainer = document.getElementById('zoneListContainer');
const clearZoneBtn = document.getElementById('clearZoneBtn');
const settingsModal = document.getElementById('settingsModal');
const settingsBtn = document.getElementById('settingsBtn');

const gradients = {
    default: { 0.4: 'blue', 0.65: 'lime', 1: 'red' },
    fire: { 0.2: 'black', 0.5: 'maroon', 0.8: 'red', 1: 'orange' },
    deepblue: { 0.2: '#000033', 0.5: '#003399', 0.8: '#0099ff', 1: '#ccffff' },
    toxic: { 0.2: 'black', 0.5: 'green', 0.8: '#ccff00', 1: 'white' },
    sunrise: { 0.2: 'black', 0.5: 'purple', 0.8: 'magenta', 1: 'gold' }
};

// --- 3. ЛОГИКА ПЕРЕТАСКИВАНИЯ ОКНА (DRAGGABLE) ---
function makeDraggable(windowId, headerId) {
    const win = document.getElementById(windowId);
    const header = document.getElementById(headerId);
    let isDragging = false, offsetX = 0, offsetY = 0;

    header.addEventListener('mousedown', (e) => {
        isDragging = true;
        offsetX = e.clientX - win.offsetLeft;
        offsetY = e.clientY - win.offsetTop;
        win.style.zIndex = 2500; // Поднимаем активное окно на передний план
    });

    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        let newLeft = e.clientX - offsetX;
        let newTop = e.clientY - offsetY;
        
        // Не даем окну улететь за границы экрана
        newLeft = Math.max(0, Math.min(newLeft, window.innerWidth - win.offsetWidth));
        newTop = Math.max(60, Math.min(newTop, window.innerHeight - win.offsetHeight)); // 60px — высота топбара

        win.style.left = `${newLeft}px`;
        win.style.top = `${newTop}px`;
        win.style.right = 'auto'; // Сбрасываем right, если он был задан в CSS
    });

    document.addEventListener('mouseup', () => { isDragging = false; });
}

// Делаем оба модальных окна перетаскиваемыми
makeDraggable('settingsModal', 'settingsHeader');
makeDraggable('zoneModal', 'zoneHeader');

// Кнопка открытия/закрытия настроек в топ-баре
settingsBtn.addEventListener('click', () => {
    const isHidden = settingsModal.style.display === 'none';
    settingsModal.style.display = isHidden ? 'block' : 'none';
});

// --- 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---
function parseExcelDate(rawDate) {
    if (rawDate === undefined || rawDate === null) return null;
    if (typeof rawDate === 'number' && rawDate > 30000 && rawDate < 60000) {
        const dateObj = new Date((rawDate - 25569) * 86400 * 1000);
        return `${String(dateObj.getDate()).padStart(2, '0')}.${String(dateObj.getMonth() + 1).padStart(2, '0')}.${dateObj.getFullYear()}`;
    }
    if (rawDate instanceof Date) {
        return `${String(rawDate.getDate()).padStart(2, '0')}.${String(rawDate.getMonth() + 1).padStart(2, '0')}.${rawDate.getFullYear()}`;
    }
    return String(rawDate).trim();
}

// --- 5. ЧТЕНИЕ ФАЙЛА И ОБРАБОТКА ---
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    uploadLabel.innerHTML = `⏳ Читання...`;
    uploadLabel.classList.remove('success');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellDates: true});
            const firstSheetName = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
            processData(jsonData, file.name);
        } catch (err) {
            console.error(err);
            statusDiv.innerHTML = "Помилка";
            uploadLabel.innerHTML = "📂 Завантажити Excel / CSV";
        }
    };
    reader.readAsArrayBuffer(file);
});

function processData(data, filename) {
    if (!data || data.length === 0) return;

    const heatData = [];
    let bounds = new L.LatLngBounds();
    activeFilters = {};
    dynamicSlicers.innerHTML = '';
    clearZone();

    const keys = Object.keys(data[0]);
    const findKey = (s) => keys.find(k => k.toLowerCase().includes(s));
    
    const latKey = findKey('lat') || findKey('шир');
    const lngKey = findKey('lng') || findKey('lon') || findKey('долг');
    
    if (!latKey || !lngKey) {
        statusDiv.innerHTML = "Відсутні координати";
        alert("Не знайдено стовпці з широтою (Lat) та довготою (Lng)");
        return;
    }

    const latIdx = keys.indexOf(latKey);
    const lngIdx = keys.indexOf(lngKey);
    const maxCoordIdx = Math.max(latIdx, lngIdx);
    const minCoordIdx = Math.min(latIdx, lngIdx);

    const valKey = keys[maxCoordIdx + 1]; // Поле сразу после координат
    const nodeKey = keys.find(k => {
        const low = k.toLowerCase();
        return low.includes('склад') || low.includes('узел') || low.includes('назв') || low.includes('name') || low.includes('пункт') || low.includes('відділ') || low.includes('отдел') || low.includes('опс') || low.includes('объект') || low.includes('обєкт') || low.includes('точк') || low.includes('мпо');
    }) || keys[0];

    const allFilterKeys = keys.filter((k, idx) => idx < minCoordIdx);

    data.forEach((row, index) => {
        const lat = parseFloat(row[latKey]);
        const lng = parseFloat(row[lngKey]);
        
        let val = 1;
        if (valKey && row[valKey] !== undefined && row[valKey] !== null && row[valKey] !== '') {
            const rawVal = row[valKey];
            if (typeof rawVal === 'number') val = rawVal;
            else {
                const parsed = parseFloat(String(rawVal).replace(',', '.'));
                if (!isNaN(parsed)) val = parsed;
            }
        }

        const nodeName = row[nodeKey] ? String(row[nodeKey]).trim() : `Точка #${index + 1}`;

        if (!isNaN(lat) && !isNaN(lng)) {
            const pt = { lat, lng, val, nodeName, filters: {} };
            
            allFilterKeys.forEach(fKey => {
                let rawVal = row[fKey];
                if (fKey.toLowerCase().includes('дат') || fKey.toLowerCase().includes('date') || rawVal instanceof Date || (typeof rawVal === 'number' && rawVal > 30000 && rawVal < 60000)) {
                    const parsedDate = parseExcelDate(rawVal);
                    if (parsedDate) rawVal = parsedDate;
                }
                pt.filters[fKey] = rawVal !== undefined && rawVal !== null && rawVal !== '' ? String(rawVal).trim() : 'Не вказано';
            });

            heatData.push(pt);
            bounds.extend([lat, lng]);
        }
    });

    if (heatData.length > 0) {
        cachedData = heatData;
        renderSlicers(allFilterKeys);
        drawHeatmap();
        map.fitBounds(bounds);
        
        uploadLabel.innerHTML = `✅ ${filename}`;
        uploadLabel.classList.add('success');
        updateStatusText();
    }
}

// --- 6. ОТРИСОВКА СРЕЗОВ ---
// --- 6. ОТРИСОВКА СРЕЗОВ (С АККОРДЕОНОМ) ---
function renderSlicers(filterKeys) {
    dynamicSlicers.innerHTML = '';
    let createdSlicersCount = 0;

    filterKeys.forEach(fKey => {
        activeFilters[fKey] = new Set();
        const uniqueVals = Array.from(new Set(cachedData.map(pt => pt.filters[fKey]))).filter(Boolean);
        if (uniqueVals.length <= 1) return;

        createdSlicersCount++;

        // Сортировка (даты — хронологически, числа — по возрастанию, текст — по алфавиту)
        if (fKey.toLowerCase().includes('дат') || fKey.toLowerCase().includes('date')) {
            uniqueVals.sort((a, b) => {
                const parseD = (s) => {
                    const p = s.split(/[\.\-\/]/);
                    return p.length === 3 ? new Date(p[2], p[1]-1, p[0]) : new Date(s);
                };
                return parseD(a) - parseD(b);
            });
        } else {
            uniqueVals.sort((a, b) => {
                const numA = parseFloat(a), numB = parseFloat(b);
                if (!isNaN(numA) && !isNaN(numB) && String(numA) === a && String(numB) === b) return numA - numB;
                return a.localeCompare(b, 'ru');
            });
        }

        // Создаем контейнер аккордеона
        const section = document.createElement('div');
        section.className = 'accordion-item';
        
        // ⚡ УМНОЕ СВОРАЧИВАНИЕ: если фильтров больше 3, 4-й и последующие сворачиваем сразу
        const isCollapsed = createdSlicersCount > 3;

        section.innerHTML = `
            <div class="accordion-header ${isCollapsed ? 'collapsed' : ''}">
                <span>🔍 ${fKey}</span>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="accordion-body ${isCollapsed ? 'collapsed' : ''}">
                <div class="slicer-container"></div>
                <div class="hint-text">Ctrl / Cmd для мультивибору</div>
            </div>
        `;
        
        // Вешаем клик на заголовок для сворачивания / разворачивания
        const header = section.querySelector('.accordion-header');
        const body = section.querySelector('.accordion-body');
        
        header.addEventListener('click', () => {
            header.classList.toggle('collapsed');
            body.classList.toggle('collapsed');
        });

        dynamicSlicers.appendChild(section);

        // Заполняем фильтр значениями
        const container = section.querySelector('.slicer-container');

        uniqueVals.forEach(val => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'slicer-btn';
            const count = cachedData.filter(pt => pt.filters[fKey] === val).length;
            btn.innerHTML = `<span>${val}</span><span style="opacity:0.5; font-size:0.75rem">${count}</span>`;
            
            btn.addEventListener('click', function(e) {
                const isMulti = e.ctrlKey || e.metaKey;
                const currentSet = activeFilters[fKey];

                if (isMulti) {
                    if (currentSet.has(val)) { currentSet.delete(val); btn.classList.remove('active'); }
                    else { currentSet.add(val); btn.classList.add('active'); }
                } else {
                    const wasOnlyThis = currentSet.has(val) && currentSet.size === 1;
                    currentSet.clear();
                    container.querySelectorAll('.slicer-btn').forEach(b => b.classList.remove('active'));
                    if (!wasOnlyThis) { currentSet.add(val); btn.classList.add('active'); }
                }
                
                drawHeatmap();
                updateStatusText();
                if (currentDrawnLayer) calculateZoneStats(currentDrawnLayer);
            });
            container.appendChild(btn);
        });
    });

    // Кнопка сброса всех фильтров
    if (createdSlicersCount > 0) {
        const resetBtn = document.createElement('button');
        resetBtn.className = 'btn';
        resetBtn.style.cssText = 'width: 100%; background: #334155; padding: 10px; margin-bottom: 10px; border-color: rgba(255,255,255,0.2);';
        resetBtn.innerHTML = '🔄 Збити усі фільтри';
        resetBtn.onclick = () => {
            for (const key in activeFilters) activeFilters[key].clear();
            dynamicSlicers.querySelectorAll('.slicer-btn').forEach(b => b.classList.remove('active'));
            drawHeatmap();
            updateStatusText();
            if (currentDrawnLayer) calculateZoneStats(currentDrawnLayer);
        };
        dynamicSlicers.insertBefore(resetBtn, dynamicSlicers.firstChild);
    }
}

function getFilteredData() {
    return cachedData.filter(pt => {
        for (const [fKey, selectedSet] of Object.entries(activeFilters)) {
            if (selectedSet.size > 0 && !selectedSet.has(pt.filters[fKey])) return false;
        }
        return true;
    });
}

function updateStatusText() {
    const total = cachedData.length;
    const filtered = getFilteredData().length;
    const hasActiveFilters = Object.values(activeFilters).some(set => set.size > 0);
    statusDiv.innerHTML = hasActiveFilters ? `${filtered} из ${total}` : `Всього: ${total}`;
}

// --- 7. ОТРИСОВКА ТЕПЛОВОЙ КАРТЫ ---
function drawHeatmap() {
    if (cachedData.length === 0) return;
    const r = parseInt(radiusInput.value);
    const b = parseInt(blurInput.value);
    const m = parseFloat(maxInput.value);
    const gKey = gradientSelect.value;

    document.getElementById('radiusVal').textContent = r;
    document.getElementById('blurVal').textContent = b;
    document.getElementById('maxVal').textContent = m;

    if (heatLayer) map.removeLayer(heatLayer);

    const options = { radius: r, blur: b, maxZoom: 10, max: m };
    if (gKey !== 'default') options.gradient = gradients[gKey];

    const leafletPoints = getFilteredData().map(pt => [pt.lat, pt.lng, pt.val]);
    heatLayer = L.heatLayer(leafletPoints, options).addTo(map);
}

// --- 8. АНАЛИЗ ВЫДЕЛЕННОЙ ЗОНЫ ---
function isPointInPolygon(lat, lng, polyCoords) {
    let x = lat, y = lng, inside = false;
    for (let i = 0, j = polyCoords.length - 1; i < polyCoords.length; j = i++) {
        let xi = polyCoords[i].lat, yi = polyCoords[i].lng;
        let xj = polyCoords[j].lat, yj = polyCoords[j].lng;
        let intersect = ((yi > y) !== (yj > y)) && (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
        if (intersect) inside = !inside;
    }
    return inside;
}

function calculateZoneStats(layer) {
    const activePoints = getFilteredData();
    let pointsInZone = [];

    if (layer instanceof L.Rectangle) {
        const bounds = layer.getBounds();
        pointsInZone = activePoints.filter(pt => bounds.contains([pt.lat, pt.lng]));
    } else if (layer instanceof L.Polygon) {
        const latlngs = layer.getLatLngs()[0];
        pointsInZone = activePoints.filter(pt => isPointInPolygon(pt.lat, pt.lng, latlngs));
    }

    let totalSum = 0;
    const nodeTotals = {};

    pointsInZone.forEach(pt => {
        totalSum += pt.val;
        if (!nodeTotals[pt.nodeName]) nodeTotals[pt.nodeName] = 0;
        nodeTotals[pt.nodeName] += pt.val;
    });

    const sortedNodes = Object.entries(nodeTotals).sort((a, b) => b[1] - a[1]);

    zoneTotalVal.textContent = `Сума: ${totalSum.toLocaleString('ru-RU', { maximumFractionDigits: 2 })}`;
    zoneListContainer.innerHTML = '';

    if (sortedNodes.length === 0) {
        zoneListContainer.innerHTML = '<div style="color:var(--text-muted); text-align:center; padding:10px;">В зоні відсутні точки</div>';
    } else {
        sortedNodes.forEach(([name, val]) => {
            const item = document.createElement('div');
            item.className = 'zone-item';
            item.innerHTML = `<span class="zone-item-name">${name}</span><span class="zone-item-val">${val.toLocaleString('ru-RU', { maximumFractionDigits: 2 })}</span>`;
            zoneListContainer.appendChild(item);
        });
    }

    zoneModal.style.display = 'block'; // Показываем модалку без затемнения фона
}

function clearZone() {
    if (currentDrawnLayer) {
        map.removeLayer(currentDrawnLayer);
        currentDrawnLayer = null;
    }
    zoneModal.style.display = 'none';
}

map.on('pm:create', function(e) {
    clearZone();
    currentDrawnLayer = e.layer;
    currentDrawnLayer.setStyle({ color: '#38bdf8', fillColor: '#38bdf8', fillOpacity: 0.2 });
    calculateZoneStats(currentDrawnLayer);
    currentDrawnLayer.on('pm:edit', () => calculateZoneStats(currentDrawnLayer));
});

map.on('pm:remove', (e) => { if (e.layer === currentDrawnLayer) clearZone(); });
clearZoneBtn.addEventListener('click', clearZone);

// --- 9. ЭКСПОРТ КАРТИНКИ ---
snapBtn.addEventListener('click', function() {
    const originalText = snapBtn.textContent;
    snapBtn.textContent = "⏳ Снимок...";
    
    screenshoter.takeScreen('blob').then(blob => {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'heatmap_export.png';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        snapBtn.textContent = originalText;
    }).catch(e => {
        console.error(e);
        alert("Ошибка при создании снимка.");
        snapBtn.textContent = originalText;
    });
});

// Слушатели настроек
radiusInput.addEventListener('input', drawHeatmap);
blurInput.addEventListener('input', drawHeatmap);
maxInput.addEventListener('input', drawHeatmap);
gradientSelect.addEventListener('change', drawHeatmap);