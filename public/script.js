document.addEventListener('DOMContentLoaded', () => {
    const fileUpload = document.getElementById('file-upload');
    const outputTableDiv = document.getElementById('output-table');
    const separationListOutputDiv = document.getElementById('separation-list-output');
    const printButton = document.getElementById('print-separation-list');

    fileUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) {
            return;
        }

        const reader = new FileReader();

        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            const { groupedOrders, separationList } = processYampiData(jsonData);

            displayData(groupedOrders, outputTableDiv);
            displaySeparationList(separationList, separationListOutputDiv);
        };

        reader.readAsArrayBuffer(file);
    });

    printButton.addEventListener('click', () => {
        printSeparationList(separationListOutputDiv);
    });

    function processYampiData(data) {
        const headerRow = data[0];
        const buyerNameColIndex = headerRow.indexOf('cliente');
        const productColIndex = headerRow.indexOf('produto');
        const quantityColIndex = headerRow.indexOf('quantidade');
        const orderNumberColIndex = headerRow.indexOf('numero_pedido');

        if (buyerNameColIndex === -1 || productColIndex === -1 || quantityColIndex === -1 || orderNumberColIndex === -1) {
            console.error('Colunas essenciais (cliente, produto, quantidade, numero_pedido) não encontradas na planilha.');
            alert('A planilha não possui todas as colunas esperadas: "cliente", "produto", "quantidade", "numero_pedido". Por favor, verifique o arquivo.');
            return { groupedOrders: {}, separationList: {} };
        }

        const groupedOrders = {};
        const consolidatedProducts = {};

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const buyerName = row[buyerNameColIndex];
            const productName = row[productColIndex];
            const quantity = parseInt(String(row[quantityColIndex]).replace(',', '.'), 10);
            const orderNumber = row[orderNumberColIndex];

            if (!buyerName || !productName || isNaN(quantity) || !orderNumber) {
                console.warn('Linha com dados inválidos ignorada:', row);
                continue;
            }

            if (!groupedOrders[buyerName]) {
                groupedOrders[buyerName] = {};
            }
            if (!groupedOrders[buyerName][orderNumber]) {
                groupedOrders[buyerName][orderNumber] = [];
            }
            groupedOrders[buyerName][orderNumber].push({
                product: productName,
                quantity: quantity
            });

            if (consolidatedProducts[productName]) {
                consolidatedProducts[productName] += quantity;
            } else {
                consolidatedProducts[productName] = quantity;
            }
        }
        return { groupedOrders: groupedOrders, separationList: consolidatedProducts };
    }

    function setupHoverEffects() {
        const orderRows = document.querySelectorAll('.sales-table tbody tr[data-order-id]');

        orderRows.forEach(row => {
            row.addEventListener('mouseover', () => {
                const orderId = row.getAttribute('data-order-id');
                const relatedRows = document.querySelectorAll(`.sales-table tbody tr[data-order-id="${orderId}"]`);
                relatedRows.forEach(relatedRow => {
                    relatedRow.classList.add('hovered-order');
                });
            });

            row.addEventListener('mouseout', () => {
                const orderId = row.getAttribute('data-order-id');
                const relatedRows = document.querySelectorAll(`.sales-table tbody tr[data-order-id="${orderId}"]`);
                relatedRows.forEach(relatedRow => {
                    relatedRow.classList.remove('hovered-order');
                });
            });
        });
    }

    function displayData(data, container) {
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
                <th>Produtos</th>
            </tr>
        `;
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        let rowNumber = 1;

        for (const buyerName in data) {
            for (const orderNumber in data[buyerName]) {
                const products = data[buyerName][orderNumber];

                const row = document.createElement('tr');
                row.setAttribute('data-order-id', orderNumber);
                row.classList.add('order-row-start');

                const numberCell = document.createElement('td');
                numberCell.textContent = rowNumber++;
                numberCell.rowSpan = products.length;
                row.appendChild(numberCell);

                const buyerCell = document.createElement('td');
                buyerCell.textContent = buyerName;
                buyerCell.rowSpan = products.length;
                row.appendChild(buyerCell);

                const orderCell = document.createElement('td');
                orderCell.textContent = orderNumber;
                orderCell.rowSpan = products.length;
                row.appendChild(orderCell);

                const productCell = document.createElement('td');
                productCell.innerHTML = `<span>${products[0].product}</span> (Qtd: <span class="product-quantity">${products[0].quantity}</span>)`;
                row.appendChild(productCell);

                tbody.appendChild(row);

                for (let i = 1; i < products.length; i++) {
                    const extraRow = document.createElement('tr');
                    extraRow.setAttribute('data-order-id', orderNumber);
                    extraRow.classList.add('order-row-follow');

                    const extraProductCell = document.createElement('td');
                    extraProductCell.innerHTML = `<span>${products[i].product}</span> (Qtd: <span class="product-quantity">${products[i].quantity}</span>)`;
                    extraRow.appendChild(extraProductCell);
                    tbody.appendChild(extraRow);
                }
            }
        }
        table.appendChild(tbody);
        container.appendChild(table);

        setupHoverEffects();
    }

    function displaySeparationList(data, container) {
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

    function printSeparationList(elementToPrint) {
        // --- INÍCIO DA CORREÇÃO ---
        // 0. Verifica se já existe um print-wrapper e o remove
        const existingPrintWrapper = document.getElementById('print-wrapper');
        if (existingPrintWrapper) {
            existingPrintWrapper.remove(); // Ou document.body.removeChild(existingPrintWrapper);
        }
        // --- FIM DA CORREÇÃO ---

        // 1. Clonar o conteúdo que queremos imprimir
        const contentToPrint = elementToPrint.cloneNode(true);

        // 2. Criar um wrapper temporário para a impressão
        const printWrapper = document.createElement('div');
        printWrapper.id = 'print-wrapper';
        
        const printTitle = document.createElement('h1');
        printTitle.textContent = 'Lista de Separação de Produtos';
        printWrapper.appendChild(printTitle);

        printWrapper.appendChild(contentToPrint);

        // 3. Adicionar o wrapper ao body temporariamente
        document.body.appendChild(printWrapper);

        // 4. Adicionar uma classe ao body para ativar os estilos de impressão
        document.body.classList.add('printing');

        // 5. Chamar a função de impressão do navegador
        window.print();

        // 6. Remover a classe do body e o wrapper após a impressão
        const afterPrintHandler = () => {
            document.body.classList.remove('printing');
            if (document.body.contains(printWrapper)) {
                printWrapper.remove(); // Melhor usar .remove() se for compatível
            }
            window.removeEventListener('afterprint', afterPrintHandler);
        };

        window.addEventListener('afterprint', afterPrintHandler);

        // Fallback para navegadores que não suportam 'afterprint' ou para remoção mais rápida
        // Reduzi o timeout, pois a remoção será garantida no início da próxima chamada.
        setTimeout(() => {
            if (document.body.classList.contains('printing')) {
                document.body.classList.remove('printing');
                if (document.body.contains(printWrapper)) {
                    printWrapper.remove();
                }
            }
        }, 300); // Um tempo menor, apenas para visualização de retorno
    }
});