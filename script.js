const monthForms = {
    1: { ok:["січень","січні"], all:["січень","січня","січні"] },
    2: { ok:["лютий","лютому"], all:["лютий","лютого","лютому"] },
    3: { ok:["березень","березні"], all:["березень","березня","березні"] },
    4: { ok:["квітень","квітні"], all:["квітень","квітня","квітні"] },
    5: { ok:["травень","травні"], all:["травень","травня","травні"] },
    6: { ok:["червень","червні"], all:["червень","червня","червні"] },
    7: { ok:["липень","липні"], all:["липень","липня","липні"] },
    8: { ok:["серпень","серпні"], all:["серпень","серпня","серпні"] },
    9: { ok:["вересень","вересні"], all:["вересень","вересня","вересні"] },
    10:{ ok:["жовтень","жовтні"], all:["жовтень","жовтня","жовтні"] },
    11:{ ok:["листопад","листопаді"], all:["листопад","листопада","листопаді"] },
    12:{ ok:["грудень","грудні"], all:["грудень","грудня","грудні"] }
};

function checkHeader(rows, month, year, messages){

    const header1 = (rows[0] || []).join(" ").toLowerCase();
    const header2 = (rows[1] || []).join(" ").toLowerCase();

    const headerText = header1 + " " + header2;

    let foundMonth = null;
    let foundMonthKey = null;

    // шукаємо будь-який місяць у тексті
    for(const key in monthForms){

        const forms = monthForms[key];

        for(const m of forms.all){

            if(headerText.includes(m)){
                foundMonth = m;
                foundMonthKey = Number(key);
                break;
            }

        }

        if(foundMonth) break;

    }

    if(!foundMonth){
        messages.push(`❌ У шапці файлу не знайдено жодного місяця`);
        return true;
    }

    // перевірка правильного місяця
    if(foundMonthKey !== month){

        const expected = monthForms[month].ok[0];

        messages.push(`❌ У шапці файлу знайдено місяць "${foundMonth}". Очікується "${expected}"`);
        return true;
    }

    // перевірка відмінку
    const correctForms = monthForms[month].ok;

    if(!correctForms.includes(foundMonth)){

        messages.push(`❌ Неправильний відмінок місяця "${foundMonth}". Очікується "${correctForms.join('" або "')}"`);
        return true;

    }

    // перевірка "місяць + рік" разом
    let correctMonthYear = false;
    
    for(const m of monthForms[month].ok){
    
        const pattern = new RegExp(`${m}\\s+${year}`);
    
        if(pattern.test(headerText)){
            correctMonthYear = true;
            break;
        }
    
    }
    
    if(!correctMonthYear){
        messages.push(`❌ У шапці файлу не знайдено "${monthForms[month].ok.join('" або "')}" ${year}`);
        return true;
    }
    return false;

}

function excelDateToJSDate(serial) {

    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;                                        
    const date_info = new Date(utc_value * 1000);

    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());

}

function parseDates(cellValue) {

    if (!cellValue) return [];

    // якщо Excel число
    if (typeof cellValue === "number") {
        return [excelDateToJSDate(cellValue)];
    }

    let raw = String(cellValue);

    const delimiters = ['\n', '\r', ',', ';', ' '];

    delimiters.forEach(d => {
        raw = raw.split(d).join('\n');
    });

    const parts = raw.split('\n').map(p => p.trim()).filter(p => p);

    const parsed = [];

    for (let p of parts) {

        let date;

        if (/^\d{2}\.\d{2}\.\d{4}$/.test(p)) {

            const [d,m,y] = p.split(".");
            date = new Date(y, m-1, d);

        } else if (/^\d{4}-\d{2}-\d{2}$/.test(p)) {

            const [y,m,d] = p.split("-");
            date = new Date(y, m-1, d);

        } else if (!isNaN(p)) {

            // Excel serial number
            date = excelDateToJSDate(Number(p));

        } else {

            return `Невірний формат дати: '${p}'`;

        }

        parsed.push(date);

    }

    return parsed;

}

function daysBetween(start, end) {

    const diff = end - start;

    return Math.floor(diff / (1000*60*60*24)) + 1;

}

function processPeriodGroup(row, idxStart, idxEnd, idxDays, rowNum, label, messages) {

    const startCell = row[idxStart];
    const endCell = row[idxEnd];
    const totalDaysCell = row[idxDays];

    if (!startCell && !endCell && !totalDaysCell) {
        return {periods:[], days:0, error:false};
    }

    const startDates = parseDates(startCell);
    const endDates = parseDates(endCell);

    if (typeof startDates === "string") {
        messages.push(`❌ [Рядок ${rowNum} | ${label}] ${startDates}`);
        return {periods:[], days:0, error:true};
    }

    if (typeof endDates === "string") {
        messages.push(`❌ [Рядок ${rowNum} | ${label}] ${endDates}`);
        return {periods:[], days:0, error:true};
    }

    if (startDates.length !== endDates.length) {
        messages.push(`❌ [Рядок ${rowNum} | ${label}] Нерівна кількість початкових і кінцевих дат.`);
        return {periods:[], days:0, error:true};
    }

    let totalCalculatedDays = 0;
    const periods = [];

    const month = Number(document.getElementById("month").value);
    const year = Number(document.getElementById("year").value);

    for (let i=0;i<startDates.length;i++) {

        const start = startDates[i];
        const end = endDates[i];

        if (
            start.getMonth()+1 !== month ||
            end.getMonth()+1 !== month ||
            start.getFullYear() !== year ||
            end.getFullYear() !== year
        ) {

            messages.push(`❌ [Рядок ${rowNum} | ${label}] Період не належить обліковому місяцю ${month}.${year}`);

            return {periods:[], days:0, error:true};
        }

        if (end < start) {
            messages.push(`❌ [Рядок ${rowNum} | ${label}] Кінець ${end.toLocaleDateString()} раніше початку ${start.toLocaleDateString()}.`);
            continue;
        }

        totalCalculatedDays += daysBetween(start,end);
        periods.push([start,end]);

    }

    const sorted = periods.slice().sort((a,b)=>a[0]-b[0]);

    for (let i=1;i<sorted.length;i++) {

        const prevEnd = sorted[i-1][1];
        const currentStart = sorted[i][0];

        if (currentStart <= prevEnd) {

            messages.push(`⚠️ [Рядок ${rowNum} | ${label}] Періоди перетинаються.`);
            return {periods, days:totalCalculatedDays, error:true};

        }

    }

    if (Number(totalDaysCell) !== totalCalculatedDays) {

        messages.push(`❌ [Рядок ${rowNum} | ${label}] Очікувалося ${totalCalculatedDays} днів, але вказано ${totalDaysCell}.`);
        return {periods, days:totalCalculatedDays, error:true};

    }

    return {periods, days:totalCalculatedDays, error:false};

}
function checkPeriods(workbook) {

    const sheetName = "30,100";

    if (!workbook.Sheets[sheetName]) {
        const availableSheets = workbook.SheetNames && workbook.SheetNames.length
            ? workbook.SheetNames.join(", ")
            : "жодного листа";
    
        return [`❌ Лист '${sheetName}' не знайдено. Доступні листи: ${availableSheets}`];
    }

    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet,{header:1});

    const messages = [];
    let anyErrors = false;
    
    const month = Number(document.getElementById("month").value);
    const year = Number(document.getElementById("year").value);
    
    if(checkHeader(rows, month, year, messages)){
        anyErrors = true;
    }

    if(checkRowNumbers(rows, messages)){
    anyErrors = true;
    }

    for (let r=5;r<rows.length;r++) {

        const row = rows[r];
        const rowNum = r+1;

        const res1 = processPeriodGroup(row,4,5,6,rowNum,"E/F/G",messages);
        const res2 = processPeriodGroup(row,8,9,10,rowNum,"I/J/K",messages);

        if (res1.error || res2.error) anyErrors = true;

        for (let p1 of res1.periods) {
            for (let p2 of res2.periods) {

                if (!(p1[1] < p2[0] || p2[1] < p1[0])) {

                    messages.push(`⚠️ [Рядок ${rowNum}] Перетин між групами E/F/G і I/J/K`);
                    anyErrors = true;

                }

            }
        }

    }

    if (!anyErrors) {
        messages.push("✅ Усі рядки коректні.");
    }

    return messages;

}

const fileInput = document.getElementById("fileInput");
const result = document.getElementById("result");
const copyBtn = document.getElementById("copyBtn");

fileInput.addEventListener("change", function(){

    if(fileInput.files.length > 0){

        document.getElementById("selectedFile").innerText =
            "📄 Обраний файл: " + fileInput.files[0].name;

        // очищаємо попередній результат
        result.innerHTML = "";

        // вимикаємо кнопку копіювання
        copyBtn.disabled = true;

    }

});

document.getElementById("checkBtn").addEventListener("click", function() {

    const file = fileInput.files[0];

    const month = Number(document.getElementById("month").value);
    const year = Number(document.getElementById("year").value);

    if (!file) {
        result.innerHTML = "❗ Оберіть Excel файл.";
        return;
    }

    if (!month || !year) {
        result.innerHTML = "❗ Вкажіть місяць і рік.";
        return;
    }

    result.innerHTML = "⏳ Перевіряю файл...";
    
    const reader = new FileReader();

    reader.onload = function(event) {

        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data,{type:"array"});

        const messages = checkPeriods(workbook);

        const errorCount = messages.filter(m => m.includes("❌") || m.includes("⚠️")).length;

        if (errorCount > 0) {
            copyBtn.disabled = false;
            result.innerHTML =
                `<b>🔴 Знайдено помилок: ${errorCount}</b><br><br>` +
                messages.join("<br>");
        } else {
            result.innerHTML = `<b>🟢 Помилок не знайдено</b>`;
            copyBtn.disabled = true;
        }

        fileInput.value = "";

        document.getElementById("selectedFile").innerText =
            "✔ Перевірка завершена. Можете обрати інший файл.";

    };

    reader.readAsArrayBuffer(file);

});

document.getElementById("copyBtn").addEventListener("click", function(){

    const text = document.getElementById("result").innerText;

    if(!text){
        alert("Немає повідомлень для копіювання.");
        return;
    }

    navigator.clipboard.writeText(text);

    alert("Повідомлення скопійовано 📋");

});

function checkRowNumbers(rows, messages){

    let startRow = null;

    for(let r = 5; r < rows.length; r++){

        const value = Number(rows[r]?.[0]);

        if(value === 1){
            startRow = r;
            break;
        }

    }

    if(startRow === null){
        messages.push("❌ Не знайдено початок нумерації (номер 1)");
        return true;
    }

    let expected = 1;
    let peopleCount = 0;
    let lastNumber = 0;
    let hasError = false;

    for(let r = startRow; r < rows.length; r++){

        const row = rows[r];

        if(
            !row[2] && !row[3] &&
            !row[4] && !row[5] && !row[6] &&
            !row[8] && !row[9] && !row[10]
        ){
            break;
        }

        peopleCount++;

        const value = row[0];

        if(value === undefined || value === null || value === ""){
            messages.push(`❌ Відсутній порядковий номер (рядок ${r+1})`);
            hasError = true;
            continue;
        }

        const number = Number(value);

        if(isNaN(number)){
            messages.push(`❌ Невірний формат номера (рядок ${r+1})`);
            hasError = true;
            continue;
        }

        if(number !== expected){
            messages.push(`❌ Порушена нумерація: очікувався ${expected}, але знайдено ${number} (рядок ${r+1})`);
            hasError = true;
        }

        lastNumber = number;
        expected++;

    }

    if(lastNumber !== peopleCount){
        messages.push(`❌ Кількість номерів (${lastNumber}) не співпадає з кількістю людей (${peopleCount})`);
        hasError = true;
    }

    return hasError;

}
