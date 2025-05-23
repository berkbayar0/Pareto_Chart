    const excludedIndexes = [1, 2, 5, 16];
    const Hatalar_list = Array.from({ length: 93 }, (_, i) => "H" + i);
    const Duruslar_list = Array.from({ length: 25 }, (_, i) => "D" + i);

    function excelDateToJSDate(serial) {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
    }

    function isValidDateFormat(str) {
      return /^\d{2}\.\d{2}\.\d{4}$/.test(str);
    }

    function isTimeStringValid(timeString) {
      if (typeof timeString === "number") return timeString >= 0 && timeString <= 1;
      if (typeof timeString === "string") {
        const timeRegex = /^(0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/;
        return timeRegex.test(timeString);
      }
      return false;
    }

    function isPatternMatch(input) {
      if (typeof input !== "string") return false;
      const cleaned = input.replace(/\s/g, '');
      return /^(\d+\*\w+(\+\d+\*\w+)*)$/.test(cleaned);
    }

    function extractTextValues(input) {
      if (!input || typeof input !== "string") return [];
      const pattern = /\b\d+\s*\*\s*(\w+)\s*\b/g;
      const matches = [];
      let match;
      while ((match = pattern.exec(input)) !== null) {
        if (match[1]) matches.push(match[1].toUpperCase());
      }
      return matches;
    }

    function checkListContains(list1, list2) {
      return list1.every(item => list2.includes(item));
    }

    function isPatternValid(text, validList) {
      if (!text || text === "0") return true;
      if (!isPatternMatch(text)) return false;
      const codes = extractTextValues(text);
      return checkListContains(codes, validList);
    }

    // Tarihe göre hafta numarası hesaplama fonksiyonu (ISO 8601 formatında)
    function getWeekFromDate(dateString) {
      if (!dateString || typeof dateString !== 'string') return '';
      
      try {
        // dd.mm.yyyy formatını parse et
        const parts = dateString.split('.');
        if (parts.length !== 3) return '';
        
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]);
        const year = parseInt(parts[2]);
        
        if (isNaN(day) || isNaN(month) || isNaN(year)) return '';
        
        const date = new Date(year, month - 1, day);
        
        // ISO 8601 hafta hesaplama
        // Yılın 4 Ocak'ı her zaman 1. haftadadır
        const jan4 = new Date(year, 0, 4);
        
        // 4 Ocak'ın hangi günde olduğunu bul (0=Pazar, 1=Pazartesi, ...)
        const jan4Day = jan4.getDay();
        
        // Yılın ilk pazartesini bul
        const firstMonday = new Date(jan4);
        firstMonday.setDate(jan4.getDate() - (jan4Day === 0 ? 6 : jan4Day - 1));
        
        // Hedef tarih ile ilk pazartesi arasındaki gün farkını hesapla
        const diffTime = date.getTime() - firstMonday.getTime();
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        
        // Hafta numarasını hesapla
        const weekNumber = Math.floor(diffDays / 7) + 1;
        
        // Eğer negatif hafta numarası çıkarsa, bir önceki yılın son haftası
        if (weekNumber <= 0) {
          const prevYear = year - 1;
          const prevYearLastWeek = getLastWeekOfYear(prevYear);
          return `${prevYear}-W${String(prevYearLastWeek).padStart(2, '0')}`;
        }
        
        // Yılın son haftasını kontrol et
        const lastWeekOfYear = getLastWeekOfYear(year);
        if (weekNumber > lastWeekOfYear) {
          return `${year + 1}-W01`;
        }
        
        return `${year}-W${String(weekNumber).padStart(2, '0')}`;
      } catch (error) {
        console.error('Hafta hesaplama hatası:', error);
        return '';
      }
    }
    
    // Yılın son hafta numarasını hesapla
    function getLastWeekOfYear(year) {
      const dec31 = new Date(year, 11, 31);
      const jan4NextYear = new Date(year + 1, 0, 4);
      const jan4NextYearDay = jan4NextYear.getDay();
      
      // Eğer 31 Aralık, bir sonraki yılın ilk haftasındaysa
      if (jan4NextYearDay <= 4) { // Perşembe veya öncesi
        const firstMondayNextYear = new Date(jan4NextYear);
        firstMondayNextYear.setDate(jan4NextYear.getDate() - (jan4NextYearDay === 0 ? 6 : jan4NextYearDay - 1));
        
        if (dec31 >= firstMondayNextYear) {
          return 1; // Aslında bir sonraki yılın ilk haftası
        }
      }
      
      // Normal hesaplama
      const jan4 = new Date(year, 0, 4);
      const jan4Day = jan4.getDay();
      const firstMonday = new Date(jan4);
      firstMonday.setDate(jan4.getDate() - (jan4Day === 0 ? 6 : jan4Day - 1));
      
      const diffTime = dec31.getTime() - firstMonday.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 1000));
      return Math.floor(diffDays / 7) + 1;
    }

    function updateProgress(current, total) {
      const percentage = Math.round((current / total) * 100);
      const progressFill = document.getElementById('progress-fill');
      const progressText = document.getElementById('progress-text');
      
      progressFill.style.width = percentage + '%';
      progressText.textContent = `${current}/${total} satır işlendi... (${percentage}%)`;
    }

    function showProgress() {
      document.getElementById('progress-container').style.display = 'block';
    }

    function hideProgress() {
      document.getElementById('progress-container').style.display = 'none';
    }

    function exportToExcel(headers, rows, fileName) {
      // İstenen sütun başlıkları ve sıralaması
      const desiredHeaders = [
        "ID",
        "Email", 
        "Tarih",
        "Hafta",
        "Proses Adı",        // "Process adı" yerine
        "OK Ürün",           // "ok ürün adeti" yerine
        "Fireler",           // "hata adet ve kodlar" yerine
        "Fire Açıklaması"
      ];
      const headerKeyMap = {
        "proses adı": "process adı",
        "ok ürün": "ok ürün adeti",
        "fireler": "hata adet ve kodlar",
        "fire açıklaması": "h0 için açıklama"
      };

      // Orijinal başlıkların indekslerini bul - daha esnek eşleştirme
      const actualIndices = {};
      headers.forEach((header, index) => {
        if (header && typeof header === 'string') {
          const cleanHeader = header.trim().toLowerCase();
          actualIndices[cleanHeader] = index;
          
          // Process adı için alternatif eşleştirmeler
          if (cleanHeader.includes('process') || cleanHeader.includes('proses')) {
            actualIndices['process adı'] = index;
          }
          if (cleanHeader.includes('ürün') && cleanHeader.includes('adet')) {
            actualIndices['ok ürün adeti'] = index;
          }
          if (cleanHeader.includes('hata') && (cleanHeader.includes('adet') || cleanHeader.includes('kod'))) {
            actualIndices['hata adet ve kodlar'] = index;
          }
          if (cleanHeader.includes('h0') || (cleanHeader.includes('açıklama') && cleanHeader.includes('h0'))) {
            actualIndices['h0 için açıklama'] = index;
          }
        }
      });

      console.log('Başlık eşleştirmeleri:', actualIndices); // Debug için

      // Her istenen başlık için orijinal indeksi bul
      const headerMap = desiredHeaders.map(desiredHeader => {
        if (desiredHeader === "Hafta") {
          return "HAFTA"; // Özel işaretçi
        }
        
        const matchKey = headerKeyMap[desiredHeader.toLowerCase()] || desiredHeader.toLowerCase();
        const foundIndex = actualIndices[matchKey];
  
        if (foundIndex !== undefined) {
            return foundIndex;
        }

        for (const [key, value] of Object.entries(actualIndices)) {
            if (key.includes(matchKey.split(' ')[0]) || matchKey.includes(key.split(' ')[0])) {
                return value;
            }
        }

        return -1;
      });

      // Filtrelenmiş satırları oluştur
      const filteredRows = rows.map(row => {
        return headerMap.map((idx, headerIndex) => {
          if (idx === "HAFTA") {
            // Tarih sütunundan hafta bilgisini hesapla
            const dateIndex = actualIndices["tarih"];
            const dateValue = dateIndex !== undefined ? row[dateIndex] : "";
            return getWeekFromDate(dateValue);
          }
          
          if (idx === -1) return "(eksik)"; // Başlık bulunamadı
          return row[idx] ?? "";
        });
      });

      // Excel workbook oluştur
      const wb = XLSX.utils.book_new();
      const wsData = [desiredHeaders, ...filteredRows];
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      
      // Kolon genişliklerini ayarla
      ws['!cols'] = desiredHeaders.map(() => ({ wch: 20 }));
      
      // Worksheet'i workbook'a ekle
      XLSX.utils.book_append_sheet(wb, ws, 'Veriler');
      
      // Dosyayı indir
      XLSX.writeFile(wb, fileName);
    }

    function createTable(headers, rows, tableId, isErrorTable = false) {
      const section = document.createElement('div');
      section.className = 'table-section';

      // Header kısmı
      const tableHeader = document.createElement('div');
      tableHeader.className = `table-header ${isErrorTable ? 'error-header' : ''}`;
      
      const titleDiv = document.createElement('div');
      titleDiv.className = 'table-title';
      
      const icon = isErrorTable ? '❌' : '✅';
      const title = isErrorTable ? 'Hatalı Kayıtlar' : 'Başarılı Kayıtlar';
      
      titleDiv.innerHTML = `
        ${icon} ${title}
        <span class="table-count">${rows.length} kayıt</span>
      `;
      
      // Kontrol butonları container
      const controlsDiv = document.createElement('div');
      controlsDiv.className = 'table-controls';
      
      // Export butonu
      const exportBtn = document.createElement('button');
      exportBtn.className = 'export-button';
      exportBtn.innerHTML = '📊 Excel\'e Aktar';
      exportBtn.onclick = () => {
        const fileName = isErrorTable ? 'Hatali_Kayitlar.xlsx' : 'Basarili_Kayitlar.xlsx';
        exportToExcel(headers, rows, fileName);
      };
      
      // Toggle butonu
      const toggleBtn = document.createElement('button');
      toggleBtn.className = 'toggle-button';
      toggleBtn.innerHTML = '<span class="toggle-icon">▼</span> Göster/Gizle';
      
      controlsDiv.appendChild(exportBtn);
      controlsDiv.appendChild(toggleBtn);
      
      tableHeader.appendChild(titleDiv);
      tableHeader.appendChild(controlsDiv);

      // Tablo wrapper
      const wrapper = document.createElement('div');
      wrapper.className = 'table-wrapper';
      
      // Tablo oluştur
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      const tbody = document.createElement('tbody');
      tbody.id = tableId + "-body";

      const headerRow = document.createElement('tr');
      headers.forEach((h, idx) => {
        if (!excludedIndexes.includes(idx)) {
          const th = document.createElement('th');
          th.textContent = h;
          headerRow.appendChild(th);
        }
      });
      thead.appendChild(headerRow);
      table.appendChild(thead);

      rows.forEach(row => {
        const tr = document.createElement('tr');
        if (isErrorTable) tr.classList.add("error");
        row.forEach((cell, idx) => {
          if (!excludedIndexes.includes(idx)) {
            const td = document.createElement('td');
            td.textContent = cell !== undefined ? cell : '';
            tr.appendChild(td);
          }
        });
        tbody.appendChild(tr);
      });

      table.appendChild(tbody);
      wrapper.appendChild(table);

      // Toggle fonksiyonu
      toggleBtn.onclick = () => {
        const isCollapsed = wrapper.classList.contains('collapsed');
        if (isCollapsed) {
          wrapper.classList.remove('collapsed');
          toggleBtn.classList.remove('collapsed');
        } else {
          wrapper.classList.add('collapsed');
          toggleBtn.classList.add('collapsed');
        }
      };

      section.appendChild(tableHeader);
      section.appendChild(wrapper);
      return section;
    }

    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const tablesContainer = document.getElementById('tables');

    // Drag and drop olayları
    dropZone.addEventListener('dragover', e => {
      e.preventDefault();
      dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', e => {
      e.preventDefault();
      dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', e => {
      e.preventDefault();
      dropZone.classList.remove('drag-over');
      handleFile(e.dataTransfer.files[0]);
    });

    // File input değişiklik olayı
    fileInput.addEventListener('change', e => {
      if (e.target.files[0]) {
        handleFile(e.target.files[0]);
      }
    });

    // Drop zone tıklama olayı - dosya input'unu temizle
    dropZone.addEventListener('click', (e) => {
      // Eğer tıklama file input'un kendisinden geliyorsa, hiçbir şey yapma
      if (e.target === fileInput) return;
      
      // File input'u temizle ve tıkla
      fileInput.value = '';
      fileInput.click();
    });

    function handleFile(file) {
      if (!file) return;

      showProgress();
      tablesContainer.innerHTML = '';

      const reader = new FileReader();
      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const headers = json[0];
        const successRows = [];
        const errorRows = [];
        let successCount = 0;
        let errorCount = 0;
        const totalRows = json.length - 1;

        // İlerleme göstergesi için batch işleme
        function processRows(startIndex) {
          const batchSize = 50;
          const endIndex = Math.min(startIndex + batchSize, json.length);

          for (let i = startIndex; i < endIndex; i++) {
            const row = json[i];
            const newRow = [...row];
            let isError = false;

            try {
              if (typeof row[6] === 'number') {
                const dateObj = excelDateToJSDate(row[6]);
                const formatted = dateObj.toLocaleDateString('tr-TR');
                newRow[6] = formatted;
                if (!isValidDateFormat(formatted)) isError = true;
              }

              [9, 10].forEach(idx => {
                if (typeof row[idx] === 'number') {
                  const hourDecimal = row[idx] * 24;
                  const hh = String(Math.floor(hourDecimal)).padStart(2, '0');
                  const mm = String(Math.round((hourDecimal % 1) * 60)).padStart(2, '0');
                  newRow[idx] = `${hh}:${mm}`;
                } else if (typeof row[idx] === 'string') {
                  newRow[idx] = row[idx];
                }
                if (!isTimeStringValid(newRow[idx])) isError = true;
              });

              if (!isPatternValid(newRow[12], Hatalar_list)) isError = true;
              if (!isPatternValid(newRow[14], Duruslar_list)) isError = true;

            } catch (err) {
              console.error(`Satır ${i + 1} hatalı:`, err);
              isError = true;
            }

            if (isError) {
              errorRows.push(newRow);
              errorCount++;
            } else {
              successRows.push(newRow);
              successCount++;
            }
          }

          updateProgress(endIndex - 1, totalRows);

          if (endIndex < json.length) {
            // Sonraki batch'i işlemek için kısa bir bekleme
            setTimeout(() => processRows(endIndex), 10);
          } else {
            // İşlem tamamlandı
            finishProcessing();
          }
        }

        function finishProcessing() {
          hideProgress();

          if (successCount > 0) {
            const mainTable = createTable(headers, successRows, "main-table");
            tablesContainer.appendChild(mainTable);
          }

          if (errorCount > 0) {
            const errTable = createTable(headers, errorRows, "error-table", true);
            tablesContainer.appendChild(errTable);
          }

          // Dosya işlemi tamamlandıktan sonra input'u temizle
          fileInput.value = '';
        }

        // İşleme başla
        if (totalRows > 0) {
          processRows(1);
        } else {
          hideProgress();
          fileInput.value = '';
        }
      };

      reader.readAsArrayBuffer(file);
    }