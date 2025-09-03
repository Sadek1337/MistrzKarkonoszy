const sheetSelect = document.getElementById('sheetSelect');
let workbookData = null;
const wynikiZawodow = "Mistrz_Karkonoszy_po_runda3.xlsx";

window.addEventListener('DOMContentLoaded', () => {
    fetch(wynikiZawodow)
        .then(res => res.arrayBuffer())
        .then(data => {
            workbookData = XLSX.read(data, { type: 'array' });
            sheetSelect.innerHTML = '';
            workbookData.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });
            renderTable();
        });
});

sheetSelect.addEventListener('change', renderTable, false);

function renderTable() {
    if (!workbookData) return;
    const sheetName = sheetSelect.value;
    const worksheet = workbookData.Sheets[sheetName];
    if (!worksheet) return;
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: 0 });
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    if (jsonData.length === 0) return;
    const headers = Object.keys(jsonData[0]);
    const trHead = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        trHead.appendChild(th);
    });
    tableHead.appendChild(trHead);
    const pozKey = headers.find(h => h.toLowerCase().includes('poz'));
    const pozMap = {};
    jsonData.forEach((row, idx) => {
        const poz = row[pozKey];
        if (!pozMap[poz]) pozMap[poz] = [];
        pozMap[poz].push(idx);
    });
    jsonData.forEach((row, idx) => {
        const tr = document.createElement('tr');
        const poz = row[pozKey];
        let rowClass = '';
        if (poz === 1) rowClass = 'winner-bg';
        else if (pozMap[poz].length > 1) rowClass = 'exaequo-bg';
        headers.forEach(h => {
            const td = document.createElement('td');
            td.textContent = row[h];
            if (rowClass) td.classList.add(rowClass);
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}
