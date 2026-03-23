// ============================
// KONFIGURACJA I STAN APLIKACJI
// ============================
// ⬇️ TUTAJ WKLEJ SWÓJ KLUCZ API W ZAMIAN ZA TEKST "TWÓJ_KLUCZ" ⬇️
const HARDCODED_API_KEY = "AIzaSyDDZ5e-d7r1VH_w3z_ppfxt9Nk_996OSIY";

let invoices = JSON.parse(localStorage.getItem('invoices') || '[]');
let monthlyBalances = JSON.parse(localStorage.getItem('monthlyBalances') || '{}');
let currentMonth = new Date().toISOString().slice(0, 7); // Format: "YYYY-MM"

// Elementy UI
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const loadingOverlay = document.getElementById('loadingOverlay');
const loadingText = document.getElementById('loadingText');
const invoicesTbody = document.getElementById('invoicesTbody');
const totalAmountElement = document.getElementById('totalAmount');
const monthSelector = document.getElementById('monthSelector');
const exportCsvBtn = document.getElementById('exportCsvBtn');

const setArchiveFolderBtn = document.getElementById('setArchiveFolderBtn');
const archiveStatus = document.getElementById('archiveStatus');
let archiveDirHandle = null;

// IndexedDB - do zapamiętania uprawnień w przeglądarce
const dbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open('SkanerFakturDB', 1);
    request.onupgradeneeded = (e) => {
        e.target.result.createObjectStore('handles');
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
});

async function saveDirHandle(handle) {
    const db = await dbPromise;
    return new Promise((resolve, reject) => {
        const tx = db.transaction('handles', 'readwrite');
        const req = tx.objectStore('handles').put(handle, 'archiveFolder');
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
    });
}

async function getDirHandle() {
    const db = await dbPromise;
    return new Promise((resolve, reject) => {
        const tx = db.transaction('handles', 'readonly');
        const req = tx.objectStore('handles').get('archiveFolder');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

async function verifyPermission(fileHandle, readWrite) {
    const options = { mode: readWrite ? 'readwrite' : 'read' };
    if ((await fileHandle.queryPermission(options)) === 'granted') return true;
    if ((await fileHandle.requestPermission(options)) === 'granted') return true;
    return false;
}

const initialBalanceInput = document.getElementById('initialBalance');
const expensesTotalElement = document.getElementById('expensesTotal');
const finalBalanceElement = document.getElementById('finalBalance');

const settingsModal = document.getElementById('settingsModal');
const apiKeyInput = document.getElementById('apiKeyInput');

const deleteModal = document.getElementById('deleteModal');
const cancelDeleteBtn = document.getElementById('cancelDeleteBtn');
const confirmDeleteBtn = document.getElementById('confirmDeleteBtn');
let invoiceIdToDelete = null;

// ============================
// INICJALIZACJA
// ============================
document.addEventListener('DOMContentLoaded', () => {
    // Sprawdź, czy mamy klucz API (zaszyty w kodzie lub zapisany w przeglądarce)
    const savedApiKey = HARDCODED_API_KEY !== "TWÓJ_KLUCZ" ? HARDCODED_API_KEY : localStorage.getItem('gemini_api_key');
    if (!savedApiKey) {
        settingsModal.classList.remove('hidden');
    } else {
        apiKeyInput.value = savedApiKey;
    }

    renderMonthOptions();
    renderTable();
});

// ============================
// OBSŁUGA PLIKÓW (DRAG & DROP)
// ============================
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

function setLoading(isLoading, text = "Przetwarzanie faktury przez AI...") {
    if (isLoading) {
        loadingText.textContent = text;
        loadingOverlay.classList.remove('hidden');
    } else {
        loadingOverlay.classList.add('hidden');
    }
}

// ============================
// PRZETWARZANIE PRZEZ AI (GEMINI)
// ============================
async function handleFile(file) {
    const apiKey = HARDCODED_API_KEY !== "TWÓJ_KLUCZ" ? HARDCODED_API_KEY : localStorage.getItem('gemini_api_key');
    if (!apiKey) {
        alert("Wymagany jest klucz API Gemini. Wpisz go w ustawieniach.");
        settingsModal.classList.remove('hidden');
        return;
    }

    if (!file.type.startsWith('image/') && file.type !== 'application/pdf') {
        alert("Wybierz plik obrazu (zdjęcie, zrzut ekranu, JPG, PNG) lub dokument PDF.");
        return;
    }

    setLoading(true);
    try {
        let base64Clean;
        let mimeType = file.type;

        if (file.type === 'application/pdf') {
            base64Clean = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result.split(',')[1]);
                reader.readAsDataURL(file);
            });
        } else {
            // Konwersja do base64 z użyciem Canvas by zoptymalizować wielkość dla AI
            const base64Data = await resizeImageToBase64(file);
            // Odcięcie prefiksu data:image/jpeg;base64,
            base64Clean = base64Data.split(',')[1];
            mimeType = 'image/jpeg';
        }
        
        // Komunikacja z API Google Gemini - najnowszy model Flash (2.5)
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                contents: [{
                    parts: [
                        {
                            text: 'Przeanalizuj tę fakturę/paragon. Twoim zadaniem jest wyciągnąć 3 informacje i zwrócić je TYLKO jako czysty obiekt JSON. Nie używaj znaczników markdown. Pola w JSON: "invoice_number" (tekst - numer faktury lub paragonu, jeśli brak wpisz "Brak numeru"), "description" (krótki, czytelny opis wydatku, np. "Paliwo Orlen", "Części komputerowe", max 3-4 słowa), "amount_total" (liczba zmiennoprzecinkowa, kropka dziesiętna, suma brutto z faktury).'
                        },
                        {
                            inlineData: {
                                mimeType: mimeType,
                                data: base64Clean
                            }
                        }
                    ]
                }],
                generationConfig: {
                    temperature: 0.1, // Niska wartość by AI było precyzyjne i powtarzalne
                    responseMimeType: "application/json"
                }
            })
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error?.message || "Błąd podczas łączenia z API Gemini.");
        }

        const data = await response.json();
        const jsonText = data.candidates[0].content.parts[0].text;
        // Zabezpieczenie na wypadek, gdyby AI mimo instrukcji dodało znaczniki Markdown
        const cleanJsonText = jsonText.replace(/```json/g, '').replace(/```/g, '').trim();
        
        let extracted;
        try {
            extracted = JSON.parse(cleanJsonText);
        } catch (e) {
            console.error("Nie udało się sparsować odpowiedzi AI:", jsonText);
            throw new Error("AI nie zwróciło poprawnego formatu.");
        }

        // Dodanie nowej faktury do zestawienia
        const newInvoice = {
            id: Date.now().toString(),
            date: new Date().toISOString().slice(0, 10), // Dziś np 2024-03-24
            month: new Date().toISOString().slice(0, 7), // Tylko miesiąc
            invoice_number: extracted.invoice_number,
            description: extracted.description,
            amount: parseFloat(extracted.amount_total) || 0
        };

        invoices.push(newInvoice);
        saveData();
        
        // Upewnijmy się, że widok jest ustawiony na aktualny miesiąc w którym dodano fakturę
        currentMonth = newInvoice.month;
        renderMonthOptions();
        renderTable();

    } catch (error) {
        console.error(error);
        alert("Wystąpił błąd: " + error.message);
    } finally {
        setLoading(false);
        // Czyszczenie inputu
        fileInput.value = '';
    }
}

// ============================
// OBSŁUGA DANYCH I INTERFEJSU
// ============================

function saveData() {
    localStorage.setItem('invoices', JSON.stringify(invoices));
}

function saveBalances() {
    localStorage.setItem('monthlyBalances', JSON.stringify(monthlyBalances));
}

initialBalanceInput.addEventListener('input', (e) => {
    monthlyBalances[currentMonth] = parseFloat(e.target.value) || 0;
    saveBalances();
    updateBalancesUI();
});

function updateBalancesUI(currentTotal = null) {
    if (currentTotal === null) {
        currentTotal = invoices.filter(inv => inv.month === currentMonth).reduce((sum, inv) => sum + inv.amount, 0);
    }
    const initial = parseFloat(monthlyBalances[currentMonth]) || 0;
    // Raport kasowy zakłada odjęcie wydatków od salda początkowego
    const final = initial - currentTotal; 
    
    expensesTotalElement.textContent = currentTotal.toFixed(2) + ' PLN';
    finalBalanceElement.textContent = final.toFixed(2) + ' PLN';
    
    if (final < 0) {
        finalBalanceElement.style.color = '#dc2626'; // Czerwony na minusie
    } else {
        finalBalanceElement.style.color = '#d92323'; // Akcentowy normalnie
    }
}

function renderMonthOptions() {
    // Zdobywamy listę unikalnych miesięcy
    const months = [...new Set(invoices.map(inv => inv.month))];
    
    // Zawsze upewnij się, że obecny miesiąc jest na liście, by nie było pusto na start
    const thisMonth = new Date().toISOString().slice(0, 7);
    if (!months.includes(thisMonth)) {
        months.push(thisMonth);
    }
    
    months.sort().reverse(); // Najnowsze na górze

    monthSelector.innerHTML = '';
    months.forEach(month => {
        const option = document.createElement('option');
        option.value = month;
        // Formatowanie RRRR-MM do przyjaznego widoku
        const [yyyy, mm] = month.split('-');
        const dateObj = new Date(yyyy, mm - 1, 1);
        const niceName = dateObj.toLocaleDateString('pl-PL', { month: 'long', year: 'numeric' });
        
        // Pierwsza litera wielka
        option.textContent = niceName.charAt(0).toUpperCase() + niceName.slice(1);
        if (month === currentMonth) option.selected = true;
        monthSelector.appendChild(option);
    });
    
    // Załaduj stan dla nowo wybranego miesiąca
    initialBalanceInput.value = (monthlyBalances[currentMonth] || 0).toFixed(2);
}

monthSelector.addEventListener('change', (e) => {
    currentMonth = e.target.value;
    initialBalanceInput.value = (monthlyBalances[currentMonth] || 0).toFixed(2);
    renderTable();
});

function renderTable() {
    invoicesTbody.innerHTML = '';
    let currentTotal = 0;

    const currentMonthInvoices = invoices.filter(inv => inv.month === currentMonth);
    
    // Sortuj najnowsze na górze
    currentMonthInvoices.sort((a, b) => b.id.localeCompare(a.id));

    if (currentMonthInvoices.length === 0) {
        invoicesTbody.innerHTML = '<tr><td colspan="5" class="empty-state">Brak faktur w tym miesiącu. Przeciągnij plik powyżej, aby dodać pierwszą.</td></tr>';
        totalAmountElement.textContent = '0.00 PLN';
        return;
    }

    currentMonthInvoices.forEach(inv => {
        currentTotal += inv.amount;
        const tr = document.createElement('tr');
        
        tr.innerHTML = `
            <td>${inv.date}</td>
            <td><strong>${inv.invoice_number}</strong></td>
            <td>${inv.description}</td>
            <td class="amount-col">${inv.amount.toFixed(2)} PLN</td>
            <td style="display: flex; gap: 5px;">
                ${inv.archiveFileName ? `<button class="delete-btn open-btn" data-filename="${inv.archiveFileName}" title="Otwórz zarchiwizowane zdjęcie (plik)">📂</button>` : ''}
                <button class="delete-btn" data-id="${inv.id}" title="Usuń fakturę">🗑</button>
            </td>
        `;
        invoicesTbody.appendChild(tr);
    });

    totalAmountElement.textContent = currentTotal.toFixed(2) + ' PLN';
    updateBalancesUI(currentTotal);
}

// Event Delegation (Najbezpieczniejsza i w 100% odporna na zacięcia metoda podpinania przycisków w HTML)
invoicesTbody.addEventListener('click', async (e) => {
    
    // ====== LOGIKA OTWIERANIA PLIKU ======
    const openBtn = e.target.closest('.open-btn');
    if (openBtn) {
        const fileName = openBtn.getAttribute('data-filename');
        if (fileName && archiveDirHandle) {
            try {
                const hasPerm = await verifyPermission(archiveDirHandle, false);
                if (hasPerm) {
                    const fileHandle = await archiveDirHandle.getFileHandle(fileName);
                    const fileObj = await fileHandle.getFile();
                    const url = URL.createObjectURL(fileObj);
                    window.open(url, '_blank');
                } else {
                    alert("Zezwól najpierw na dostęp do Archiwum na górze strony.");
                }
            } catch (err) {
                console.error("Błąd doczytywania pliku:", err);
                alert("Nie mogliśmy załadować tego skanu z Twojego dysku.\nMoże plik " + fileName + " został już przeniesiony lub usunięty z folderu?");
            }
        } else {
            alert("Podepnij folder Archiwum najpierw (przycisk na górze przy strefie wrzutowej).");
        }
        return;
    }


    // ====== LOGIKA KOSZA ======
    const btn = e.target.closest('.delete-btn');
    if (!btn) return; // Kliknięto w stół, ale nie w żaden z dwóch przycisków
    
    // Zapisanie wybranego ID i wywołanie autorskiego okna (omija blokadę confirm() w Chrome na file:///)
    invoiceIdToDelete = btn.getAttribute('data-id');
    deleteModal.classList.remove('hidden');
});

// Obsługa autorskiego okienka potwierdzenia usunięcia
cancelDeleteBtn.addEventListener('click', () => {
    deleteModal.classList.add('hidden');
    invoiceIdToDelete = null;
});

confirmDeleteBtn.addEventListener('click', () => {
    if (invoiceIdToDelete) {
        invoices = invoices.filter(item => String(item.id) !== String(invoiceIdToDelete));
        saveData();
        renderMonthOptions(); 
        renderTable();
    }
    deleteModal.classList.add('hidden');
    invoiceIdToDelete = null;
});

// ============================
// ARCHIWIZACJA LOKALNA plików na Dysk C: (System Access API)
// ============================
if (setArchiveFolderBtn) {
    setArchiveFolderBtn.addEventListener('click', async (e) => {
        e.stopPropagation(); // Nie wywołuj okienka wyboru pliku pod spodem w dropZone
        e.preventDefault();
        
        try {
            archiveDirHandle = await window.showDirectoryPicker({
                id: 'archiwumSkanerAI',
                mode: 'readwrite'
            });
            await saveDirHandle(archiveDirHandle); // Zapis do stałej pamięci przeglądarki!
            
            archiveStatus.textContent = `Archiwum powiązane trwale z: ${archiveDirHandle.name} ✅`;
            archiveStatus.style.color = 'var(--accent-color)';
            archiveStatus.style.fontWeight = 'bold';
        } catch (err) {
            console.log("Anulowano wybór folderu do Archiwum.", err);
        }
    });
}

// Przywracanie folderu po odświeżeniu/włączeniu strony!
window.addEventListener('load', async () => {
    try {
        const handle = await getDirHandle();
        if (handle) {
            archiveDirHandle = handle;
            archiveStatus.textContent = `Oczekuję odblokowania dostępu do spiętego folderu: ${handle.name}... kliknij dowolne skanowanie by powrócić!`;
            archiveStatus.style.color = '#FFA500';
            
            // Prośba o twardy powrót uprawnień jest wymagana przez Windowsa przy pierwszym zapisie (np podczas klikniecia rzutu pliku)
        }
    } catch(err) { console.error("Błąd przywracania folderu", err) }
});

// ============================
// PRZETWARZANIE PLIKU I AI
// ============================

function handleFile(file) {
    if (!file) return;

    let savedFileName = null;
    const safeDate = new Date().toISOString().slice(0, 10);
    const safeName = `SKAN_${safeDate}_${file.name}`;

    // ----- ZAPIS DO FIZYCZNEGO FOLDERU (W TLE) -----
    if (archiveDirHandle) {
        savedFileName = safeName; // Zapisz by użyć jej wokół zapytań o AI
        (async () => {
            try {
                // Odnowienie/Wymuszenie zgody jeśli przeglądarka pamięta uchwyt
                const hasPerm = await verifyPermission(archiveDirHandle, true);
                if (hasPerm) {
                    archiveStatus.textContent = `Archiwum trwale i bezpiecznie wpięte do: ${archiveDirHandle.name} ✅`;
                    archiveStatus.style.color = 'var(--accent-color)';
                    
                    const fileHandle = await archiveDirHandle.getFileHandle(safeName, { create: true });
                    const writable = await fileHandle.createWritable();
                    await writable.write(file);
                    await writable.close();
                    console.log("Plik został autozapisany!");
                }
            } catch (err) {
                console.error("Błąd archiwizacji pliku na dysku:", err);
            }
        })();
    }
    // ------------------------------------------------

    // Aktualizacja UI na ładowanie
    loadingOverlay.classList.remove('hidden');
    
    // Konwersja na Base64 i zrzut do AI (Gemini 2.5 Flash)
    const apiKey = HARDCODED_API_KEY !== "TWÓJ_KLUCZ" ? HARDCODED_API_KEY : localStorage.getItem('gemini_api_key');
    if (!apiKey) {
        alert("Brak klucza API! Skonfiguruj go w Ustawieniach.");
        loadingOverlay.classList.add('hidden');
        settingsModal.classList.remove('hidden');
        return;
    }

    const processBase64 = async (base64Data, mimeType) => {
        try {
            const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{
                        parts: [
                            { text: "Przeanalizuj poniższą fakturę lub paragon. Zwróć TYLKO czysty obiekt JSON (bez bloków markdown) o strukturze: {\"invoice_number\": \"...\", \"description\": \"Krótki opis operacji\", \"amount_total\": \"123.45\"}. Jako kwotę wstaw same cyfry z kropką." },
                            {
                                inline_data: {
                                    mime_type: mimeType,
                                    data: base64Data
                                }
                            }
                        ]
                    }],
                    generationConfig: {
                        temperature: 0.1,
                        response_mime_type: "application/json"
                    }
                })
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error?.message || "Błąd połączenia z API Google");
            }

            const data = await response.json();
            let textRaw = data.candidates[0].content.parts[0].text;
            
            // Czyszczenie ze znaczników markdown
            textRaw = textRaw.replace(/```json/g, '').replace(/```/g, '').trim();
            const extracted = JSON.parse(textRaw);

            // Oznaczanie dokładną datą i godziną
            const now = new Date();
            const yyyy = now.getFullYear();
            const mmL = String(now.getMonth() + 1).padStart(2, '0');
            const ddL = String(now.getDate()).padStart(2, '0');
            const hh = String(now.getHours()).padStart(2, '0');
            const min = String(now.getMinutes()).padStart(2, '0');
            const fullDateStr = `${yyyy}-${mmL}-${ddL} ${hh}:${min}`;

            const newInvoice = {
                id: Date.now().toString(),
                date: fullDateStr,
                month: currentMonth,
                invoice_number: extracted.invoice_number || "BRAK",
                description: extracted.description || "Brak opisu",
                amount: parseFloat(extracted.amount_total) || 0,
                archiveFileName: savedFileName
            };

            invoices.push(newInvoice);
            saveData();
            renderMonthOptions();
            renderTable();
            
        } catch (err) {
            console.error("Błąd AI:", err);
            alert("Błąd podczas analizy przez AI:\n" + err.message);
        } finally {
            loadingOverlay.classList.add('hidden');
            fileInput.value = '';
        }
    };

    if (file.type === 'application/pdf') {
        const reader = new FileReader();
        reader.onload = (e) => {
            const base64Data = e.target.result.split(',')[1];
            processBase64(base64Data, file.type);
        };
        reader.readAsDataURL(file);
    } else {
        // Wszystkie zdjęcia JPG/PNG przechodzą najpierw przez kompresję
        resizeImageToBase64(file).then(dataUrl => {
            const base64Data = dataUrl.split(',')[1];
            processBase64(base64Data, 'image/jpeg');
        }).catch(err => {
            console.error("Błąd kompresji obrazu wywołał awarię:", err);
            loadingOverlay.classList.add('hidden');
            alert("Przeglądarka ma problem z przetrawieniem tego typu zdjęcia.");
        });
    }
}

// ============================
// USTAWIENIA MODAL
// ============================
document.getElementById('settingsBtn').addEventListener('click', () => {
    settingsModal.classList.remove('hidden');
});

document.getElementById('closeSettingsBtn').addEventListener('click', () => {
    // Pozwól zamknąć tylko jeśli mamy klucz, no chyba że na siłę (ale aplikacja nie zadziała)
    settingsModal.classList.add('hidden');
});

document.getElementById('saveSettingsBtn').addEventListener('click', () => {
    const val = apiKeyInput.value.trim();
    if (val) {
        localStorage.setItem('gemini_api_key', val);
        settingsModal.classList.add('hidden');
        alert("Klucz API został zapisany bezpiecznie w Twojej przeglądarce!");
    } else {
        alert("Podaj poprawny klucz.");
    }
});

// ============================
// EKSPORT DO EXCELA / GOOGLE SHEETS W PEŁNYM FORMACIE WIZUALNYM (.XLSX)
// ============================
exportCsvBtn.addEventListener('click', async () => {
    const currentMonthInvoices = invoices.filter(inv => inv.month === currentMonth);
    if (currentMonthInvoices.length === 0) {
        alert("Brak faktur do pobrania w tym miesiącu.");
        return;
    }

    const initial = parseFloat(monthlyBalances[currentMonth]) || 0;
    const total = currentMonthInvoices.reduce((sum, inv) => sum + inv.amount, 0);
    const finalBalance = initial - total;

    // Tworzenie profesjonalnego pliku za pomocą biblioteki ExcelJS
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Skaner Faktur AI';
    const worksheet = workbook.addWorksheet('Raport Kasowy');

    // Kolumny i ich szerokości dla czytelności z ramkami
    worksheet.columns = [
        { header: '', key: 'lp', width: 6 },         // A: Lp.
        { header: '', key: 'date', width: 15 },      // B: Data
        { header: '', key: 'num', width: 22 },       // C: Numer Faktury
        { header: '', key: 'desc', width: 45 },      // D: Opis operacji
        { header: '', key: 'income', width: 18 },    // E: Przychody
        { header: '', key: 'expense', width: 18 }    // F: Rozchody
    ];

    // Nagłówek Raportu (pogrubiony tytuł na środku)
    worksheet.mergeCells('A1:F1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `ZBIORCZY RAPORT KASOWY ZA OKRES: ${currentMonth}`;
    titleCell.font = { name: 'Arial', size: 14, bold: true };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(1).height = 30;

    // Komórki Salda początkowego
    const initialLabelCell = worksheet.getCell('D3');
    initialLabelCell.value = 'STAN SALDA Z POPRZEDNIEGO MIESIĄCA:';
    initialLabelCell.font = { bold: true };
    initialLabelCell.alignment = { horizontal: 'right' };
    
    const initialValCell = worksheet.getCell('F3');
    initialValCell.value = initial;
    initialValCell.numFmt = '#,##0.00 [$$PLN]';
    initialValCell.font = { bold: true };

    // Grube, eleganckie nagłówki dla tabeli (wiersz 5)
    const headerRow = worksheet.getRow(5);
    headerRow.values = ['Lp.', 'Data Wpisu', 'Numer Dokumentu', 'Zeskanowana Treść Operacji', 'Przychód', 'Rozchód'];
    headerRow.height = 25;
    
    const headerBorder = { top: {style:'medium'}, left: {style:'medium'}, bottom: {style:'medium'}, right: {style:'medium'} };
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
        const cell = headerRow.getCell(col);
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD92323' } }; // Kolor Covebo Red w nagłówku
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = headerBorder;
    });

    // Pętla do wstrzykiwania zeskanowanych faktur (od wiersza 6)
    let rowIndex = 6;
    currentMonthInvoices.forEach((inv, i) => {
        const row = worksheet.getRow(rowIndex);
        row.values = {
            lp: i + 1,
            date: inv.date,
            num: inv.invoice_number,
            desc: inv.description,
            income: '', // puste (zakładamy koszty z faktur)
            expense: inv.amount
        };

        // Standardowe, delikatne ramki cienkie dla komórek tabeli by wyglądała profesjonalnie
        const rowBorder = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => {
            const cell = row.getCell(col);
            cell.border = rowBorder;
            cell.alignment = { vertical: 'middle', wrapText: true };
        });

        // Wyrównanie kolumn finansowych i formatowanie księgowe
        row.getCell('A').alignment = { horizontal: 'center' };
        row.getCell('B').alignment = { horizontal: 'center' };
        row.getCell('E').numFmt = '#,##0.00';
        row.getCell('F').numFmt = '#,##0.00';
        
        rowIndex++;
    });

    // Puste miejsce pod tabelą przed sumami
    rowIndex += 1;

    // Sumy i saldo na samym końcu
    const sumsLabelCell = worksheet.getCell(`D${rowIndex}`);
    sumsLabelCell.value = 'CAŁKOWITA SUMA WYDATKÓW W TABELI:';
    sumsLabelCell.font = { bold: true };
    sumsLabelCell.alignment = { horizontal: 'right' };
    
    const sumsValCell = worksheet.getCell(`F${rowIndex}`);
    sumsValCell.value = total;
    sumsValCell.numFmt = '#,##0.00 [$$PLN]';
    sumsValCell.font = { bold: true, color: { argb: 'FFD92323' } }; // Na czerwono

    rowIndex += 1;
    const endLabelCell = worksheet.getCell(`D${rowIndex}`);
    endLabelCell.value = 'STAN SALDA KOŃCOWEGO:';
    endLabelCell.font = { bold: true, size: 12 };
    endLabelCell.alignment = { horizontal: 'right' };
    
    const endValCell = worksheet.getCell(`F${rowIndex}`);
    endValCell.value = finalBalance;
    endValCell.numFmt = '#,##0.00 [$$PLN]';
    endValCell.font = { bold: true, size: 12 };
    
    if (finalBalance < 0) {
       endValCell.font.color = { argb: 'FFFF0000' };
    }

    // Ekstrakcja pliku ustrukturyzowanego i eleganckiego do pobrania z bufora
    const excelBuffer = await workbook.xlsx.writeBuffer();

    // 1. GŁÓWNA ŚCIEŻKA: Zapis automatyczny w zapamiętanym folderze bez otwierania uciążliwego wyskakującego okienka! (O ile wybrano folder archiwyzcji)
    if (archiveDirHandle) {
        try {
            const hasPerm = await verifyPermission(archiveDirHandle, true);
            if (hasPerm) {
                const safeName = `RaportKasowy_${currentMonth}.xlsx`;
                const fileHandle = await archiveDirHandle.getFileHandle(safeName, { create: true });
                const writable = await fileHandle.createWritable();
                await writable.write(excelBuffer);
                await writable.close();
                alert(`Rewelacja! Prawde mówiąc nic nie musisz zapisywać!\n\nKompletny zrobiony Raport Kasowy w formacie Excel po cichu wylądował automatycznie do wyznaczonego stałego folderu archiwum pod nazwą:\n\n${safeName}`);
                return; // Sukces!! Pomiń domyślne pobieranie.
            }
        } catch (e) {
            console.error("Cichy automat do zapisu raportu zwiódł... Wracam do okienka.", e);
        }
    }

    // 2. ŚCIEŻKA ALTERNATYWNA: Przywracamy nasz ultra-skuteczny system natywnego okienka "Zapisz Jako" (Zapis Windowsowy) gdy folder jednak nie zapamiętany.
    if (window.showSaveFilePicker) {
        try {
            const handle = await window.showSaveFilePicker({
                suggestedName: `RaportKasowy_${currentMonth}.xlsx`,
                types: [{
                    description: 'Dokument programu Excel (Uporządkowany XSLX)',
                    accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']},
                }],
            });
            const writable = await handle.createWritable();
            await writable.write(excelBuffer);
            await writable.close();
            return; // Sukces
        } catch (err) {
            if (err.name === 'AbortError') return; 
        }
    }

    // Bezpieczny fallback (gdy Chrome odmówi, lub na FireFox)
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `RaportKasowy_${currentMonth}.xlsx`;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    setTimeout(() => { document.body.removeChild(link); URL.revokeObjectURL(url); }, 100);
});

// ============================
// NARZĘDZIA POMOCNICZE
// ============================

// Funkcja optymalizująca duże zdjęcia fotograficzne. Generuje wirtualne Canvas, pomniejsza by oszczędzić 
// portfel przesyłu API (Gemini ma limity) i zwraca czyste zrównoważone zdjęcie jako Base64.
function resizeImageToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
                try {
                    // TWORZENIE CANVASU W PAMIĘCI (odporne na usunięte tagi HTML!)
                    const canvas = document.createElement('canvas');
                    let width = img.width;
                    let height = img.height;
                    const MAX_SIZE = 1600;

                    if (width > height && width > MAX_SIZE) {
                        height *= MAX_SIZE / width;
                        width = MAX_SIZE;
                    } else if (height > MAX_SIZE) {
                        width *= MAX_SIZE / height;
                        height = MAX_SIZE;
                    }

                    canvas.width = width;
                    canvas.height = height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, width, height);
                    
                    // Kompresja na w locie do JPG z jakością 80% (perfekcyjny kompromis ostrość/waga)
                    resolve(canvas.toDataURL('image/jpeg', 0.8));
                } catch(err) {
                    reject("Krytyczny błąd podczas renderowania obrazu w tle: " + err.message);
                }
            };
            img.onerror = () => reject("Błąd przetwarzania pliku jako obrazu w przeglądarce.");
            img.src = e.target.result;
        };
        reader.onerror = () => reject("Błąd odczytu surowego pliku!");
        reader.readAsDataURL(file);
    });
}
