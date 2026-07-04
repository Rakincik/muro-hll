async function startGroupAutomationScraper() {
    function showDualCopyPopup(lessonsText, usersText) {
        const container = document.createElement('div');
        container.style.position = 'fixed';
        container.style.top = '0';
        container.style.left = '0';
        container.style.width = '100%';
        container.style.height = '100%';
        container.style.backgroundColor = 'rgba(0, 0, 0, 0.8)';
        container.style.zIndex = '99999';
        container.style.display = 'flex';
        container.style.alignItems = 'center';
        container.style.justifyContent = 'center';
        container.style.fontFamily = 'sans-serif';

        const panel = document.createElement('div');
        panel.style.backgroundColor = 'white';
        panel.style.padding = '20px';
        panel.style.borderRadius = '10px';
        panel.style.width = '85%';
        panel.style.height = '80%';
        panel.style.display = 'flex';
        panel.style.flexDirection = 'column';
        panel.style.gap = '15px';

        const title = document.createElement('h3');
        title.innerText = '🎉 Otomasyon Tamamlandı! Tüm veriler hazır.';
        title.style.color = '#2ed573';
        title.style.margin = '0';
        panel.appendChild(title);

        const contentRow = document.createElement('div');
        contentRow.style.display = 'flex';
        contentRow.style.flex = '1';
        contentRow.style.gap = '20px';

        // Dersler
        const lessonsCol = document.createElement('div');
        lessonsCol.style.flex = '1';
        lessonsCol.style.display = 'flex';
        lessonsCol.style.flexDirection = 'column';
        lessonsCol.style.gap = '5px';
        const lTitle = document.createElement('h4');
        lTitle.innerText = '1. Tüm Grupların Dersleri (Kopyalamak için tıkla)';
        lTitle.style.margin = '0';
        const lTextarea = document.createElement('textarea');
        lTextarea.value = lessonsText;
        lTextarea.style.flex = '1';
        lTextarea.style.width = '100%';
        lTextarea.style.fontFamily = 'monospace';
        lTextarea.style.fontSize = '11px';
        lTextarea.style.padding = '10px';
        lTextarea.onclick = () => { lTextarea.focus(); lTextarea.select(); };
        lessonsCol.appendChild(lTitle);
        lessonsCol.appendChild(lTextarea);

        // Üyeler
        const usersCol = document.createElement('div');
        usersCol.style.flex = '1';
        usersCol.style.display = 'flex';
        usersCol.style.flexDirection = 'column';
        usersCol.style.gap = '5px';
        const uTitle = document.createElement('h4');
        uTitle.innerText = '2. Tüm Grupların Üyeleri (Kopyalamak için tıkla)';
        uTitle.style.margin = '0';
        const uTextarea = document.createElement('textarea');
        uTextarea.value = usersText;
        uTextarea.style.flex = '1';
        uTextarea.style.width = '100%';
        uTextarea.style.fontFamily = 'monospace';
        uTextarea.style.fontSize = '11px';
        uTextarea.style.padding = '10px';
        uTextarea.onclick = () => { uTextarea.focus(); uTextarea.select(); };
        usersCol.appendChild(uTitle);
        usersCol.appendChild(uTextarea);

        contentRow.appendChild(lessonsCol);
        contentRow.appendChild(usersCol);
        panel.appendChild(contentRow);

        const footer = document.createElement('div');
        footer.style.display = 'flex';
        footer.style.justifyContent = 'space-between';
        footer.style.alignItems = 'center';

        const info = document.createElement('span');
        info.innerText = 'Kutuların içine tıklayarak tümünü seçebilir ve Ctrl + C ile Excel\'e aktarabilirsiniz.';
        info.style.fontSize = '12px';
        info.style.color = '#7f8c8d';
        footer.appendChild(info);

        const closeBtn = document.createElement('button');
        closeBtn.innerText = 'Kapat';
        closeBtn.style.padding = '10px 20px';
        closeBtn.style.backgroundColor = '#ff4757';
        closeBtn.style.color = 'white';
        closeBtn.style.border = 'none';
        closeBtn.style.borderRadius = '5px';
        closeBtn.style.cursor = 'pointer';
        closeBtn.style.fontWeight = 'bold';
        closeBtn.onclick = () => {
            document.body.removeChild(container);
        };
        footer.appendChild(closeBtn);
        panel.appendChild(footer);

        container.appendChild(panel);
        document.body.appendChild(container);
    }

    if (!window.jQuery) {
        console.log("jQuery bulunamadı.");
        return;
    }

    const treeEl = jQuery('.jstree, [role="tree"], #jstree, [class*="jstree"]');
    if (treeEl.length === 0) {
        console.log("jsTree yapısı bulunamadı.");
        return;
    }

    const instance = treeEl.jstree(true);
    if (!instance) {
        console.log("jsTree instance alınamadı.");
        return;
    }

    console.log("Grup ağacı genişletiliyor...");
    instance.open_all();
    await new Promise(resolve => setTimeout(resolve, 2000));

    const flatNodes = instance.get_json('#', { flat: true });
    const targetNodes = flatNodes.filter(node => node.text !== 'Gruplar');

    console.log(`Otomasyon başladı. Toplam ${targetNodes.length} kategori/grup taranacak...`);

    const allLessons = [];
    const allUsers = [];

    // Tablonun AJAX yüklenmesinin bitmesini bekleyen ultra sabırlı fonksiyon (10 Saniye Beklemeli)
    async function waitForTableToLoad(table) {
        const maxWait = 10000; // Maksimum 10 saniye bekler
        const interval = 100;
        let waited = 0;
        
        while (waited < maxWait) {
            const wrapper = table.closest('.dataTables_wrapper');
            const processing = wrapper ? wrapper.querySelector('.dataTables_processing') : null;
            const isProcessing = processing && processing.style.display !== 'none';
            
            const rows = table.querySelectorAll('tbody tr');
            const firstRow = rows[0];
            const firstRowText = firstRow ? firstRow.innerText.trim().toLowerCase() : '';
            
            // Eğer tablo boşsa veya hâlâ "yükleniyor" uyarısı varsa beklemeye devam et
            const isStillLoading = isProcessing || 
                                   firstRowText.includes('yukleniyor') || 
                                   firstRowText.includes('loading') || 
                                   firstRowText.includes('isleniyor') || 
                                   firstRowText.includes('processing') ||
                                   firstRowText === ''; // Tamamen boşsa da yükleniyordur
                                   
            if (!isStillLoading) {
                // Yüklenme bitti (veri geldi veya "Veri bulunamadı" yazdı)
                break;
            }
            await new Promise(resolve => setTimeout(resolve, interval));
            waited += interval;
        }
        // Sayfanın kendine gelmesi için ekstra 300ms tolerans payı
        await new Promise(resolve => setTimeout(resolve, 300));
    }

    // Dinamik sütun eşlemeli tablo kazıma fonksiyonu
    function scrapeTableDynamic(table, node) {
        const tabHeaders = [];
        let headerCells = table.querySelectorAll('thead th, thead td');
        if (headerCells.length === 0) {
            const firstRow = table.querySelector('tr');
            if (firstRow) headerCells = firstRow.querySelectorAll('th, td');
        }

        headerCells.forEach((cell, idx) => {
            let text = cell.innerText.trim()
                .toLowerCase()
                .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u').replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c')
                .replace(/[^a-z0-9\s]/g, '')
                .trim()
                .replace(/\s+/g, '_');
            tabHeaders.push(text || `col_${idx}`);
        });

        // Üst grup (Parent) bilgisini al
        let parentId = node.parent && node.parent !== '#' ? node.parent : '';
        let parentName = '';
        if (parentId) {
            const pNode = instance.get_node(parentId);
            if (pNode) {
                if (pNode.text === 'Gruplar') {
                    parentId = '';
                    parentName = '';
                } else {
                    parentName = pNode.text;
                }
            }
        }

        const rows = table.querySelectorAll('tbody tr');
        const data = [];
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            // "Gösterilecek veri yok" da olabiliyor Okinar'da
            if (cells.length > 0 && cells[0].innerText.trim() !== 'Veri bulunamadı' && !cells[0].innerText.trim().toLowerCase().includes('veri yok')) {
                const rowData = {
                    grup_id: node.id,
                    grup_adi: node.text,
                    grup_ust_id: parentId, // YENİ EKLENDİ
                    grup_ust_adi: parentName // YENİ EKLENDİ
                };
                cells.forEach((cell, idx) => {
                    const key = tabHeaders[idx] || `col_${idx}`;
                    const isCheckbox = (key === 'col_0' || key === 'column_0') && cell.querySelector('input[type="checkbox"]');
                    const isAction = key.includes('islem') || key.includes('action') || key.includes('sil') || key.includes('duzenle');
                    
                    if (!isCheckbox && !isAction && key !== '') {
                        rowData[key] = cell.innerText.trim();
                    }
                });
                data.push(rowData);
            }
        });
        return data;
    }

    function findNextButtonForTable(table) {
        const pane = table.closest('.tab-pane');
        if (!pane) return null;

        let nextBtn = pane.querySelector('.paginate_button.next:not(.disabled) a, .page-item.next:not(.disabled) a');
        if (nextBtn) return nextBtn;
        
        nextBtn = pane.querySelector('.paginate_button.next:not(.disabled), .page-item.next:not(.disabled)');
        if (nextBtn) return nextBtn;

        const candidates = Array.from(pane.querySelectorAll('button, a, li, span, div[role="button"]'));
        for (let el of candidates) {
            const text = el.innerText.trim().toLowerCase();
            const isDisabled = el.hasAttribute('disabled') || 
                               el.classList.contains('disabled') || 
                               el.getAttribute('aria-disabled') === 'true' ||
                               el.closest('.disabled') !== null;

            if (isDisabled) continue;

            if (text === 'sonraki' || text === 'ileri' || text === 'next' || text === '>' || text === '»' || text === '›') {
                return el;
            }
        }
        return null;
    }

    function convertToExcelTSV(data) {
        if (data.length === 0) return '';
        
        const allKeysSet = new Set();
        data.forEach(obj => {
            Object.keys(obj).forEach(key => allKeysSet.add(key));
        });
        const headers = Array.from(allKeysSet);
        
        const headerRow = headers.join('\t');
        const dataRows = data.map(row => 
            headers.map(header => {
                let cellValue = row[header] || '';
                return cellValue.toString().replace(/\t/g, ' ').replace(/\n/g, ' ');
            }).join('\t')
        );
        return [headerRow, ...dataRows].join('\n');
    }

    // Modal popup gösterimi için script çalışmaya başladığını bildir
    console.log("Sayfa kazıma döngüsüne giriliyor...");
    for (let i = 0; i < targetNodes.length; i++) {
        const node = targetNodes[i];
        console.log(`[${i + 1}/${targetNodes.length}] Taranıyor: ${node.text} (ID: ${node.id})`);

        const anchor = document.getElementById(node.id + '_anchor');
        if (anchor) {
            anchor.click();
            // Sol ağaç tıklandıktan sonra sağ panelin yüklenmesi için 2 saniye bekle
            await new Promise(resolve => setTimeout(resolve, 2000));

            // ---- DERSLERİ KAZI ----
            const tabDersler = document.querySelector('a[href*="tab_3"], [href*="tab_3"]');
            if (tabDersler) {
                tabDersler.click();
                await new Promise(resolve => setTimeout(resolve, 800));

                const lessonsTable = document.querySelector('.tab-pane.active table');
                if (lessonsTable) {
                    await waitForTableToLoad(lessonsTable);

                    const select = lessonsTable.closest('.tab-pane')?.querySelector('select');
                    if (select && select.value !== '100' && select.value !== '-1') {
                        select.value = select.options[select.options.length - 1].value;
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        await new Promise(resolve => setTimeout(resolve, 1200));
                        await waitForTableToLoad(lessonsTable);
                    }

                    let lessonPageCount = 1;
                    let lastLessonPageFirstRowText = "";

                    while (true) {
                        const rows = lessonsTable.querySelectorAll('tbody tr');
                        if (rows.length === 0) break;

                        const currentFirst = rows[0].innerText.trim();
                        if (currentFirst === lastLessonPageFirstRowText) break;
                        lastLessonPageFirstRowText = currentFirst;

                        console.log(`  └─ Ders Sayfası ${lessonPageCount} taranıyor...`);
                        
                        const pageData = scrapeTableDynamic(lessonsTable, node);
                        allLessons.push(...pageData);

                        const nextBtn = findNextButtonForTable(lessonsTable);
                        if (!nextBtn) break;

                        nextBtn.click();
                        await new Promise(resolve => setTimeout(resolve, 1200));
                        await waitForTableToLoad(lessonsTable);
                        lessonPageCount++;
                    }
                }
            }

            // ---- KULLANICILARI KAZI ----
            const tabKullanicilar = document.querySelector('a[href*="tab_4"], [href*="tab_4"]');
            if (tabKullanicilar) {
                tabKullanicilar.click();
                await new Promise(resolve => setTimeout(resolve, 800));

                const usersTable = document.querySelector('.tab-pane.active table');
                if (usersTable) {
                    await waitForTableToLoad(usersTable);

                    const select = usersTable.closest('.tab-pane')?.querySelector('select');
                    if (select && select.value !== '100' && select.value !== '-1') {
                        select.value = select.options[select.options.length - 1].value;
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                        await new Promise(resolve => setTimeout(resolve, 1200));
                        await waitForTableToLoad(usersTable);
                    }

                    let userPageCount = 1;
                    let lastUserPageFirstRowText = "";

                    while (true) {
                        const rows = usersTable.querySelectorAll('tbody tr');
                        if (rows.length === 0) break;

                        const currentFirst = rows[0].innerText.trim();
                        if (currentFirst === lastUserPageFirstRowText) break;
                        lastUserPageFirstRowText = currentFirst;

                        console.log(`  └─ Öğrenci Sayfası ${userPageCount} taranıyor...`);
                        
                        const pageData = scrapeTableDynamic(usersTable, node);
                        allUsers.push(...pageData);

                        const nextBtn = findNextButtonForTable(usersTable);
                        if (!nextBtn) break;

                        nextBtn.click();
                        await new Promise(resolve => setTimeout(resolve, 1200));
                        await waitForTableToLoad(usersTable);
                        userPageCount++;
                    }
                }
            }
        }
    }

    console.log("Tarama bitti! Excel dönüşümleri hazırlanıyor...");

    const lessonsTsv = convertToExcelTSV(allLessons);
    const usersTsv = convertToExcelTSV(allUsers);

    showDualCopyPopup(lessonsTsv, usersTsv);
}

startGroupAutomationScraper();
