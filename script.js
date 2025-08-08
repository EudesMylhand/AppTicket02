        document.addEventListener('DOMContentLoaded', function() {
            // Récupération des éléments DOM
            const fileInput = document.getElementById('fileInput');
            const printButton = document.getElementById('printButton');
            const pagesContainer = document.getElementById('pagesContainer');
            const applySettingsBtn = document.getElementById('applySettings');
            const zoomLevel = document.getElementById('zoomLevel');
            const zoomValue = document.getElementById('zoomValue');
            const dropZone = document.getElementById('dropZone');
            
            // Paramètres configurables
            const paddingValue = document.getElementById('paddingValue');
            const paddingUnit = document.getElementById('paddingUnit');
            const cellHeight = document.getElementById('cellHeight');
            const cellHeightUnit = document.getElementById('cellHeightUnit');
            const fontSize = document.getElementById('fontSize');
            const fontSizeUnit = document.getElementById('fontSizeUnit');
            const tablePosition = document.getElementById('tablePosition');
            const textAlign = document.getElementById('textAlign');
            const marginX = document.getElementById('marginX');
            const marginXUnit = document.getElementById('marginXUnit');
            const marginY = document.getElementById('marginY');
            const marginYUnit = document.getElementById('marginYUnit');
            const cellWidth = document.getElementById('cellWidth');
            const cellWidthUnit = document.getElementById('cellWidthUnit');
            
            let currentData = null;
            let currentHeaders = null;
            
            // Initialisation des événements
            function initEvents() {
                fileInput.addEventListener('change', handleFileUpload);
                printButton.addEventListener('click', printDocument);
                applySettingsBtn.addEventListener('click', applySettings);
                zoomLevel.addEventListener('input', updateZoom);
                
                // Drag and drop
                dropZone.addEventListener('dragover', handleDragOver);
                dropZone.addEventListener('dragleave', handleDragLeave);
                dropZone.addEventListener('drop', handleDrop);
                
                // Gestion des changements de paramètres
                paddingValue.addEventListener('change', updateCSSVariables);
                paddingUnit.addEventListener('change', updateCSSVariables);
                cellHeight.addEventListener('change', updateCSSVariables);
                cellHeightUnit.addEventListener('change', updateCSSVariables);
                fontSize.addEventListener('change', updateCSSVariables);
                fontSizeUnit.addEventListener('change', updateCSSVariables);
                textAlign.addEventListener('change', updateCSSVariables);
                marginX.addEventListener('change', updateCSSVariables);
                marginXUnit.addEventListener('change', updateCSSVariables);
                marginY.addEventListener('change', updateCSSVariables);
                marginYUnit.addEventListener('change', updateCSSVariables);
                cellWidth.addEventListener('change', updateCSSVariables);
                cellWidthUnit.addEventListener('change', updateCSSVariables);
                tablePosition.addEventListener('change', updateCSSVariables);
                
                // Gestion de l'impression
                if (window.matchMedia) {
                    const mediaQueryList = window.matchMedia('print');
                    mediaQueryList.addEventListener('change', handlePrintMediaChange);
                }
            }
            
            function handleDragOver(e) {
                e.preventDefault();
                dropZone.classList.add('highlight');
            }
            
            function handleDragLeave() {
                dropZone.classList.remove('highlight');
            }
            
            function handleDrop(e) {
                e.preventDefault();
                dropZone.classList.remove('highlight');
                
                if (e.dataTransfer.files.length) {
                    fileInput.files = e.dataTransfer.files;
                    handleFileUpload({ target: fileInput });
                }
            }
            
            function handlePrintMediaChange(mql) {
                if (mql.matches) {
                    beforePrint();
                } else {
                    afterPrint();
                }
            }
            
            function updateCSSVariables() {
                document.documentElement.style.setProperty('--global-padding', `${paddingValue.value}${paddingUnit.value}`);
                document.documentElement.style.setProperty('--margin-x', `${marginX.value}${marginXUnit.value}`);
                document.documentElement.style.setProperty('--margin-y', `${marginY.value}${marginYUnit.value}`);
                document.documentElement.style.setProperty('--cell-height', `${cellHeight.value}${cellHeightUnit.value}`);
                document.documentElement.style.setProperty('--cell-width', `${cellWidth.value}${cellWidthUnit.value}`);
                document.documentElement.style.setProperty('--font-size', `${fontSize.value}${fontSizeUnit.value}`);
                document.documentElement.style.setProperty('--text-align', textAlign.value);
                
                // Mise à jour de la position de la table
                const position = tablePosition.value;
                const style = document.querySelector('style[data-dynamic-style]');
                if (style) {
                    style.textContent = `
                        .page-table-container {
                            justify-content: ${position === 'center' ? 'center' : position === 'left' ? 'flex-start' : 'flex-end'};
                        }
                    `;
                }
            }
            
            function updateZoom() {
                const zoom = zoomLevel.value;
                zoomValue.textContent = `${zoom}%`;
                pagesContainer.style.zoom = `${zoom}%`;
            }
            
            function applySettings() {
                if (currentData && currentHeaders) {
                    generatePages(currentHeaders, currentData);
                } else {
                    alert('Veuillez d\'abord charger un fichier.');
                }
            }
            
            function handleFileUpload(event) {
                resetPreview();
                
                const file = event.target.files[0];
                if (!file) {
                    return;
                }
                
                const fileExtension = file.name.split('.').pop().toLowerCase();
                
                if (fileExtension === 'csv' || fileExtension === 'txt') {
                    processCSVFile(file);
                } else if (fileExtension === 'xls' || fileExtension === 'xlsx') {
                    processExcelFile(file);
                } else {
                    alert('Format de fichier non supporté. Veuillez utiliser un fichier CSV, TXT, XLS ou XLSX.');
                }
            }

            function resetPreview() {
                currentData = null;
                currentHeaders = null;
                pagesContainer.innerHTML = '';
            }
            
            function processCSVFile(file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        processCSV(e.target.result);
                    } catch (error) {
                        alert('Erreur lors du traitement du fichier CSV: ' + error.message);
                    }
                };
                reader.onerror = function() {
                    alert('Erreur lors de la lecture du fichier.');
                };
                reader.readAsText(file);
            }
            
            function processExcelFile(file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        processExcelData(jsonData);
                    } catch (error) {
                        alert('Erreur lors du traitement du fichier Excel: ' + error.message);
                    }
                };
                reader.onerror = function() {
                    alert('Erreur lors de la lecture du fichier.');
                };
                reader.readAsArrayBuffer(file);
            }
            
            function processCSV(csvContent) {
                const lines = csvContent.split('\n').filter(line => line.trim() !== '');
                const dataLines = lines.filter(line => !line.startsWith('>'));
                
                if (dataLines.length === 0) {
                    alert('Aucune donnée trouvée dans le fichier.');
                    return;
                }
                
                const separator = detectSeparator(dataLines[0]);
                const headers = dataLines[0].split(separator)
                    .map(header => header.trim())
                    .filter(header => header !== '');
                
                const data = [];
                for (let i = 1; i < dataLines.length; i++) {
                    const values = dataLines[i].split(separator)
                        .map(value => value.trim() + ' ')
                        .filter(value => value !== ' ');
                    
                    if (values.length === headers.length) {
                        const entry = {};
                        headers.forEach((header, index) => {
                            entry[header] = values[index];
                        });
                        data.push(entry);
                    }
                }
                
                if (data.length === 0) {
                    alert('Aucune donnée valide trouvée dans le fichier.');
                    return;
                }
                
                currentHeaders = headers;
                currentData = data;
                generatePages(headers, data);
            }
            
            function detectSeparator(line) {
                const separators = [',', ';', '|', '\t'];
                let maxCount = 0;
                let detectedSeparator = ',';
                
                for (const sep of separators) {
                    const count = line.split(sep).length - 1;
                    if (count > maxCount) {
                        maxCount = count;
                        detectedSeparator = sep;
                    }
                }
                
                return detectedSeparator;
            }
            
            function processExcelData(excelData) {
                if (excelData.length === 0) {
                    alert('Aucune donnée trouvée dans le fichier.');
                    return;
                }
                
                const headers = excelData[0].map(header => header.toString().trim());
                const data = [];

                for (let i = 1; i < excelData.length; i++) {
                    const values = excelData[i];
                    if (!values || values.length === 0) continue;
                    
                    const entry = {};
                    headers.forEach((header, index) => {
                        entry[header] = values[index] !== undefined ? values[index].toString().trim() + ' ' : '';
                    });
                    data.push(entry);
                }
                
                if (data.length === 0) {
                    alert('Aucune donnée valide trouvée dans le fichier.');
                    return;
                }
                
                currentHeaders = headers;
                currentData = data;
                generatePages(headers, data);
            }
            
            function generatePages(headers, data) {
                pagesContainer.innerHTML = '';
                
                if (!data || data.length === 0) {
                    pagesContainer.innerHTML = '<p>Aucune donnée à afficher.</p>';
                    return;
                }
                
                // Création du style dynamique pour la position de la table
                const oldStyle = document.querySelector('style[data-dynamic-style]');
                if (oldStyle) oldStyle.remove();
                
                const style = document.createElement('style');
                style.setAttribute('data-dynamic-style', 'true');
                const position = tablePosition.value;
                style.textContent = `
                    .page-table-container {
                        justify-content: ${position === 'center' ? 'center' : position === 'left' ? 'flex-start' : 'flex-end'};
                    }
                `;
                document.head.appendChild(style);
                
                // Mise à jour des variables CSS
                updateCSSVariables();
                
                const itemsPerPage = 13 * 5;
                const totalPages = Math.ceil(data.length / itemsPerPage);
                
                // Génération des pages
                for (let pageNum = 0; pageNum < totalPages; pageNum++) {
                    const page = document.createElement('div');
                    page.className = 'page';
                    
                    const pageContent = document.createElement('div');
                    pageContent.className = 'page-content';
                    
                    const tableContainer = document.createElement('div');
                    tableContainer.className = 'page-table-container';
                    
                    const startIdx = pageNum * itemsPerPage;
                    const endIdx = Math.min(startIdx + itemsPerPage, data.length);
                    const pageData = data.slice(startIdx, endIdx);
                    
                    const table = document.createElement('table');
                    table.className = 'fixed-table';
                    
                    let dataIndex = 0;
                    
                    for (let i = 0; i < 13; i++) {
                        const row = document.createElement('tr');
                        
                        for (let j = 0; j < 5; j++) {
                            const cell = document.createElement('td');
                            
                            if (dataIndex < pageData.length) {
                                const item = pageData[dataIndex];
                                let cellContent = '';
                                headers.forEach(header => {
                                    if (item[header]) {
                                        cellContent += `<div><strong>${header}:</strong><br>${item[header]}</div>`;
                                    }
                                });
                                cell.innerHTML = cellContent;
                                dataIndex++;
                            } else {
                                cell.innerHTML = '';
                            }
                            row.appendChild(cell);
                        }
                        
                        table.appendChild(row);
                    }
                    
                    tableContainer.appendChild(table);
                    pageContent.appendChild(tableContainer);
                    page.appendChild(pageContent);
                    
                    // Ajout du numéro de page
                    const pageNumber = document.createElement('div');
                    pageNumber.className = 'page-number';
                    pageNumber.textContent = `Page ${pageNum + 1}/${totalPages}`;
                    page.appendChild(pageNumber);
                    
                    pagesContainer.appendChild(page);
                }
            }
            
            function printDocument() {
                if (!currentData) {
                    alert('Veuillez d\'abord charger un fichier.');
                    return;
                }
                
                // Afficher les numéros de page avant l'impression
                document.querySelectorAll('.page-number').forEach(el => {
                    el.style.display = 'block';
                });
                
                window.print();
                
                // Masquer les numéros de page après l'impression
                setTimeout(() => {
                    document.querySelectorAll('.page-number').forEach(el => {
                        el.style.display = 'none';
                    });
                }, 1000);
            }

            function beforePrint() {
                document.body.style.margin = '0';
                document.body.style.padding = '0';
            }
            
            function afterPrint() {
                document.body.style.margin = '';
                document.body.style.padding = '';
            }
            
            // Initialisation
            initEvents();
            resetPreview();
        });