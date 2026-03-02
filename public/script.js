document.addEventListener('DOMContentLoaded', () => {
    /**
     * OrderViewer Application
     * Encapsulates logic for processing, displaying, and printing order data.
     * Uses a class-based design pattern for better state management and modularity.
     */
    class OrderViewer {
        constructor() {
            // Cache DOM elements
            this.dom = {
                fileUpload: document.getElementById('file-upload'),
                outputTable: document.getElementById('output-table'),
                separationListOutput: document.getElementById('separation-list-output'),
                printSeparationBtn: document.getElementById('print-separation-list'),
                printOrdersBtn: null // Will be created dynamically
            };

            // Application State
            this.state = {
                groupedOrders: {},
                separationList: {}
            };

            this.init();
        }

        init() {
            this.injectStyles();
            this.createUI();
            this.bindEvents();
        }

        /**
         * Injects CSS for printing and UI enhancements.
         */
        injectStyles() {
            const style = document.createElement('style');
            style.textContent = `
                /* UI Enhancements */
                .hovered-order { background-color: #f0f8ff; transition: background-color 0.2s; }
                .btn-print { margin-left: 10px; }
                .shipment-icon { width: 40px; height: auto; display: block; margin: 0 auto; }

                /* Print Specific Styles */
                @media print {
                    /* Table Print Styles */
                    .sales-table { 
                        width: 100%; 
                        border-collapse: collapse; 
                        margin-top: 10px; 
                        font-size: 11px;
                    }
                    
                    .sales-table th { 
                        background-color: #eee !important; 
                        font-weight: bold;
                        border: 1px solid #000;
                        padding: 8px;
                        text-align: left;
                        -webkit-print-color-adjust: exact;
                    }

                    .sales-table td { 
                        border: 1px solid #000; 
                        padding: 6px 8px; 
                        text-align: left; 
                        vertical-align: middle;
                    }

                    /* Visibly separate orders with a thicker top border */
                    tr.order-row-start td {
                        border-top: 2px solid #000 !important;
                    }

                    /* Clean up badges for print */
                    .product-quantity {
                        color: #000 !important;
                        border: 1px solid #ccc;
                        padding: 2px 4px;
                    }

                    .product-quantity.qty-high {
                        background-color: #ff8888 !important;
                        color: #000 !important;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }

                    /* Avoid breaking rows awkwardly */
                    tr { page-break-inside: avoid; }
                    thead { display: table-header-group; }

                    /* List Print Styles */
                    .separation-ul { list-style: none; padding: 0; }
                    .separation-ul li { 
                        border-bottom: 1px solid #ccc; 
                        padding: 5px 0; 
                        display: flex; 
                        justify-content: space-between; 
                    }
                }
            `;
            document.head.appendChild(style);
        }

        /**
         * Creates additional UI elements like the Print Orders button.
         */
        createUI() {
            if (this.dom.printSeparationBtn) {
                const btn = document.createElement('button');
                btn.id = 'print-orders-list';
                btn.textContent = 'Imprimir Lista de Pedidos';
                btn.type = 'button';
                // Inherit classes from existing button for consistency, add custom class
                btn.className = (this.dom.printSeparationBtn.className || '') + ' btn-print';
                
                // Insert after the existing print button
                this.dom.printSeparationBtn.parentNode.insertBefore(btn, this.dom.printSeparationBtn.nextSibling);
                this.dom.printOrdersBtn = btn;
            }
        }

        bindEvents() {
            if (this.dom.fileUpload) {
                this.dom.fileUpload.addEventListener('change', (e) => this.handleFileUpload(e));
            }

            if (this.dom.printSeparationBtn) {
                this.dom.printSeparationBtn.addEventListener('click', () => {
                    this.printElement(this.dom.separationListOutput, 'Lista de Separação de Produtos');
                });
            }

            if (this.dom.printOrdersBtn) {
                this.dom.printOrdersBtn.addEventListener('click', () => {
                    this.printElement(this.dom.outputTable, 'Lista de Pedidos Completa');
                });
            }
        }

        handleFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    const processed = this.processData(jsonData);
                    if (processed) {
                        this.state = processed;
                        this.render();
                    }
                } catch (error) {
                    console.error("Erro ao processar arquivo:", error);
                    alert("Ocorreu um erro ao ler o arquivo. Verifique se é um Excel válido.");
                }
            };
            reader.readAsArrayBuffer(file);
        }

        processData(data) {
            const headerRow = data[0];
            // Normalize headers to lowercase for robust matching
            const headers = headerRow.map(h => String(h).toLowerCase().trim());
            
            const buyerNameColIndex = headers.indexOf('cliente');
            const productColIndex = headers.indexOf('produto');
            const quantityColIndex = headers.indexOf('quantidade');
            const orderNumberColIndex = headers.indexOf('numero_pedido');
            const shipmentColIndex = headers.indexOf('entrega');

            if (buyerNameColIndex === -1 || productColIndex === -1 || quantityColIndex === -1 || orderNumberColIndex === -1) {
                console.error('Colunas essenciais não encontradas.');
                alert('A planilha não possui todas as colunas esperadas: "cliente", "produto", "quantidade", "numero_pedido".');
                return null;
            }

            const groupedOrders = {};
            const consolidatedProducts = {};

            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (!row || row.length === 0) continue;

                const buyerName = row[buyerNameColIndex];
                const productName = row[productColIndex];
                const rawQty = row[quantityColIndex];
                const quantity = parseInt(String(rawQty || '0').replace(',', '.'), 10);
                const orderNumber = row[orderNumberColIndex];
                const shipment = shipmentColIndex !== -1 ? row[shipmentColIndex] : null;

                if (!buyerName || !productName || isNaN(quantity) || !orderNumber) {
                    continue;
                }

                // Group by Buyer -> Order Number
                if (!groupedOrders[buyerName]) groupedOrders[buyerName] = {};
                if (!groupedOrders[buyerName][orderNumber]) groupedOrders[buyerName][orderNumber] = [];
                
                groupedOrders[buyerName][orderNumber].push({
                    product: productName,
                    quantity: quantity,
                    shipment: shipment
                });

                // Consolidate for Separation List
                if (consolidatedProducts[productName]) {
                    consolidatedProducts[productName] += quantity;
                } else {
                    consolidatedProducts[productName] = quantity;
                }
            }
            return { groupedOrders, separationList: consolidatedProducts };
        }

        getShipmentDetails(shipment) {
            if (!shipment) return null;
            const s = String(shipment).trim();
            
            if (s === 'FRENET_LOGGI_LOG_DRPOFF') return { src: 'media/LOGGI.png', title: 'Loggi Dropoff' };
            if (s === 'FRENET_SEDEX_03220') return { src: 'media/SEDEX.png', title: 'Correios SEDEX' };
            if (s === 'FRENET_PAC_03298') return { src: 'media/PAC.png', title: 'Correios PAC' };
            return null;
        }

        render() {
            this.renderOrdersTable(this.state.groupedOrders, this.dom.outputTable);
            this.renderSeparationList(this.state.separationList, this.dom.separationListOutput);
            this.setupHoverEffects();
        }

        renderOrdersTable(data, container) {
            container.innerHTML = '';
            if (Object.keys(data).length === 0) {
                container.innerHTML = '<p>Nenhum dado processado para exibir.</p>';
                return;
            }

            const table = document.createElement('table');
            table.classList.add('sales-table');

            const thead = document.createElement('thead');
            thead.innerHTML = `
                <tr>
                    <th>N°</th>
                    <th>Comprador</th>
                    <th>Número do Pedido</th>
                    <th>Envio</th>
                    <th>Produtos</th>
                    <th>Qtd</th>
                </tr>
            `;
            table.appendChild(thead);

            const tbody = document.createElement('tbody');
            let rowNumber = 1;

            for (const buyerName in data) {
                for (const orderNumber in data[buyerName]) {
                    const products = data[buyerName][orderNumber];
                    const rowSpan = products.length;

                    // First row of the order
                    const firstRow = document.createElement('tr');
                    firstRow.setAttribute('data-order-id', orderNumber);
                    firstRow.classList.add('order-row-start');

                    // Helper to create cells
                    const addCell = (row, text, span = 1, isHtml = false) => {
                        const td = document.createElement('td');
                        if (span > 1) td.rowSpan = span;
                        if (isHtml) td.innerHTML = text; else td.textContent = text;
                        row.appendChild(td);
                    };

                    addCell(firstRow, rowNumber++, rowSpan);
                    addCell(firstRow, buyerName, rowSpan);
                    addCell(firstRow, orderNumber, rowSpan);

                    const shipmentInfo = this.getShipmentDetails(products[0].shipment);
                    if (shipmentInfo) {
                        addCell(firstRow, `<img src="${shipmentInfo.src}" title="${shipmentInfo.title}" class="shipment-icon" alt="${shipmentInfo.title}">`, rowSpan, true);
                    } else {
                        addCell(firstRow, '', rowSpan);
                    }
                    
                    // First product details
                    addCell(firstRow, `<span>${products[0].product}</span>`, 1, true);
                    const firstQtyClass = products[0].quantity > 1 ? 'product-quantity qty-high' : 'product-quantity';
                    addCell(firstRow, `<span class="${firstQtyClass}">${products[0].quantity}</span>`, 1, true);
                    
                    tbody.appendChild(firstRow);

                    // Subsequent products for the same order
                    for (let i = 1; i < products.length; i++) {
                        const extraRow = document.createElement('tr');
                        extraRow.setAttribute('data-order-id', orderNumber);
                        extraRow.classList.add('order-row-follow');

                        addCell(extraRow, `<span>${products[i].product}</span>`, 1, true);
                        const qtyClass = products[i].quantity > 1 ? 'product-quantity qty-high' : 'product-quantity';
                        addCell(extraRow, `<span class="${qtyClass}">${products[i].quantity}</span>`, 1, true);
                        
                        tbody.appendChild(extraRow);
                    }
                }   
            }
            table.appendChild(tbody);
            container.appendChild(table);
        }

        renderSeparationList(data, container) {
            container.innerHTML = '';
            if (Object.keys(data).length === 0) {
                container.innerHTML = '<p>Nenhum produto para a lista de separação.</p>';
                return;
            }

            const ul = document.createElement('ul');
            ul.classList.add('separation-ul');

            const sortedProducts = Object.entries(data).sort((a, b) => a[0].localeCompare(b[0]));

            sortedProducts.forEach(([productName, totalQuantity]) => {
                const li = document.createElement('li');
                li.innerHTML = `
                    <span class="product-name-list">${productName}</span> 
                    <span class="product-quantity-list">${totalQuantity}</span>
                `;
                ul.appendChild(li);
            });

            container.appendChild(ul);
        }

        setupHoverEffects() {
            const orderRows = document.querySelectorAll('.sales-table tbody tr[data-order-id]');
            orderRows.forEach(row => {
                row.addEventListener('mouseover', () => {
                    const orderId = row.getAttribute('data-order-id');
                    document.querySelectorAll(`.sales-table tbody tr[data-order-id="${orderId}"]`)
                        .forEach(r => r.classList.add('hovered-order'));
                });

                row.addEventListener('mouseout', () => {
                    const orderId = row.getAttribute('data-order-id');
                    document.querySelectorAll(`.sales-table tbody tr[data-order-id="${orderId}"]`)
                        .forEach(r => r.classList.remove('hovered-order'));
                });
            });
        }

        printElement(elementToPrint, titleText) {
            // 1. Clean up any existing print wrapper
            const existingPrintWrapper = document.getElementById('print-wrapper');
            if (existingPrintWrapper) {
                existingPrintWrapper.remove();
            }

            if (!elementToPrint || elementToPrint.innerHTML.trim() === '') {
                alert('Não há dados para imprimir.');
                return;
            }

            // 2. Create wrapper
            const printWrapper = document.createElement('div');
            printWrapper.id = 'print-wrapper';
            
            const printTitle = document.createElement('h1');
            printTitle.textContent = titleText;
            printWrapper.appendChild(printTitle);

            // 3. Clone content
            printWrapper.appendChild(elementToPrint.cloneNode(true));

            // 4. Append to body
            document.body.appendChild(printWrapper);
            document.body.classList.add('printing');

            // 5. Print
            window.print();

            // 6. Cleanup
            const afterPrintHandler = () => {
                document.body.classList.remove('printing');
                if (document.body.contains(printWrapper)) {
                    printWrapper.remove();
                }
                window.removeEventListener('afterprint', afterPrintHandler);
            };

            window.addEventListener('afterprint', afterPrintHandler);

            // Fallback for browsers that might not trigger afterprint reliably
            setTimeout(() => {
                if (document.body.classList.contains('printing')) {
                    afterPrintHandler();
                }
            }, 1000);
        }
    }

    // Initialize the application
    new OrderViewer();
});