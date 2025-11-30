// Sistema de Gestão de Planilhas - Arquivo Principal
// Versão sem módulos ES6 para funcionar direto no navegador

// Estado Global
let currentSpreadsheet = null;
let spreadsheets = [];
let versionHistory = [];

// Tipos de dados suportados
const DATA_TYPES = {
    TEXT: 'text',
    NUMBER: 'number',
    DATE: 'date',
    EMAIL: 'email',
    CURRENCY: 'currency',
    PERCENTAGE: 'percentage',
    FORMULA: 'formula'
};

// Templates predefinidos
const TEMPLATES = {
    vendas: {
        name: 'Controle de Vendas',
        description: 'Planilha para acompanhamento de vendas mensais',
        columns: [
            { name: 'Data', type: DATA_TYPES.DATE, required: true },
            { name: 'Produto', type: DATA_TYPES.TEXT, required: true },
            { name: 'Quantidade', type: DATA_TYPES.NUMBER, required: true },
            { name: 'Valor Unitário', type: DATA_TYPES.CURRENCY, required: true },
            { name: 'Total', type: DATA_TYPES.FORMULA, formula: '=C*D', required: false },
            { name: 'Vendedor', type: DATA_TYPES.TEXT, required: true }
        ],
        rows: 15,
        sampleData: [
            ['2024-01-15', 'Notebook', '5', '2500.00', '', 'João Silva'],
            ['2024-01-16', 'Mouse', '20', '45.00', '', 'Maria Santos'],
            ['2024-01-17', 'Teclado', '15', '120.00', '', 'João Silva']
        ]
    },
    orcamento: {
        name: 'Orçamento Pessoal',
        description: 'Controle de receitas e despesas',
        columns: [
            { name: 'Data', type: DATA_TYPES.DATE, required: true },
            { name: 'Descrição', type: DATA_TYPES.TEXT, required: true },
            { name: 'Categoria', type: DATA_TYPES.TEXT, required: true },
            { name: 'Tipo', type: DATA_TYPES.TEXT, required: true },
            { name: 'Valor', type: DATA_TYPES.CURRENCY, required: true }
        ],
        rows: 20,
        sampleData: [
            ['2024-01-01', 'Salário', 'Receita', 'Receita', '5000.00'],
            ['2024-01-05', 'Aluguel', 'Moradia', 'Despesa', '1500.00'],
            ['2024-01-10', 'Supermercado', 'Alimentação', 'Despesa', '600.00']
        ]
    },
    estoque: {
        name: 'Controle de Estoque',
        description: 'Gestão de estoque de produtos',
        columns: [
            { name: 'Código', type: DATA_TYPES.TEXT, required: true },
            { name: 'Produto', type: DATA_TYPES.TEXT, required: true },
            { name: 'Estoque Inicial', type: DATA_TYPES.NUMBER, required: true },
            { name: 'Entradas', type: DATA_TYPES.NUMBER, required: false },
            { name: 'Saídas', type: DATA_TYPES.NUMBER, required: false },
            { name: 'Estoque Atual', type: DATA_TYPES.FORMULA, formula: '=C+D-E', required: false },
            { name: 'Valor Unitário', type: DATA_TYPES.CURRENCY, required: true }
        ],
        rows: 12,
        sampleData: [
            ['001', 'Caneta Azul', '100', '50', '30', '', '2.50'],
            ['002', 'Caderno 200 folhas', '50', '20', '15', '', '15.00'],
            ['003', 'Lápis HB', '200', '100', '80', '', '1.20']
        ]
    },
    funcionarios: {
        name: 'Cadastro de Funcionários',
        description: 'Registro de dados de colaboradores',
        columns: [
            { name: 'Matrícula', type: DATA_TYPES.TEXT, required: true },
            { name: 'Nome', type: DATA_TYPES.TEXT, required: true },
            { name: 'Email', type: DATA_TYPES.EMAIL, required: true },
            { name: 'Cargo', type: DATA_TYPES.TEXT, required: true },
            { name: 'Salário', type: DATA_TYPES.CURRENCY, required: true },
            { name: 'Data Admissão', type: DATA_TYPES.DATE, required: true }
        ],
        rows: 10,
        sampleData: [
            ['0001', 'Ana Paula Santos', 'ana.santos@empresa.com', 'Gerente', '8000.00', '2020-03-15'],
            ['0002', 'Carlos Eduardo Lima', 'carlos.lima@empresa.com', 'Analista', '4500.00', '2021-06-20']
        ]
    },
    tarefas: {
        name: 'Lista de Tarefas',
        description: 'Gerenciamento de projetos e tarefas',
        columns: [
            { name: 'Tarefa', type: DATA_TYPES.TEXT, required: true },
            { name: 'Responsável', type: DATA_TYPES.TEXT, required: true },
            { name: 'Prazo', type: DATA_TYPES.DATE, required: true },
            { name: 'Status', type: DATA_TYPES.TEXT, required: true },
            { name: 'Prioridade', type: DATA_TYPES.TEXT, required: true },
            { name: 'Progresso', type: DATA_TYPES.PERCENTAGE, required: false }
        ],
        rows: 15,
        sampleData: [
            ['Desenvolver módulo de login', 'João', '2024-02-15', 'Em andamento', 'Alta', '60'],
            ['Revisar documentação', 'Maria', '2024-02-10', 'Pendente', 'Média', '0'],
            ['Teste de integração', 'Carlos', '2024-02-20', 'Não iniciado', 'Alta', '0']
        ]
    },
    clientes: {
        name: 'Base de Clientes',
        description: 'Gerenciamento de contatos e clientes',
        columns: [
            { name: 'ID', type: DATA_TYPES.TEXT, required: true },
            { name: 'Nome', type: DATA_TYPES.TEXT, required: true },
            { name: 'Email', type: DATA_TYPES.EMAIL, required: true },
            { name: 'Telefone', type: DATA_TYPES.TEXT, required: false },
            { name: 'Cidade', type: DATA_TYPES.TEXT, required: false },
            { name: 'Última Compra', type: DATA_TYPES.DATE, required: false }
        ],
        rows: 12,
        sampleData: [
            ['C001', 'Empresa ABC Ltda', 'contato@abc.com', '(11) 99999-9999', 'São Paulo', '2024-01-10'],
            ['C002', 'Tech Solutions', 'vendas@tech.com', '(21) 88888-8888', 'Rio de Janeiro', '2024-01-15']
        ]
    }
};

// ===========================
// INTERFACE DE LOGIN
// ===========================

function showLoginInterface() {
    document.getElementById('login-screen').classList.remove('hidden');
    document.getElementById('main-app').classList.add('hidden');
}

function showUserInterface(user) {
    document.getElementById('login-screen').classList.add('hidden');
    document.getElementById('main-app').classList.remove('hidden');

    document.getElementById('user-photo').src = user.photoURL || 'https://via.placeholder.com/40';
    document.getElementById('user-name').textContent = user.displayName || 'Usuário';
    document.getElementById('user-email').textContent = user.email || '';
}

// Handlers globais
window.handleLogin = async function() {
    try {
        await loginWithGoogle();
    } catch (error) {
        console.error('Erro no login:', error);
    }
};

window.handleLogout = async function() {
    if (confirm('Deseja realmente sair?')) {
        await logout();
    }
};

// ===========================
// INICIALIZAÇÃO
// ===========================

document.addEventListener('DOMContentLoaded', function() {
    initializeTabs();
    initializeUploadArea();
});

function initializeTabs() {
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const tabName = button.dataset.tab;

            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));

            button.classList.add('active');
            document.getElementById(tabName).classList.add('active');

            if (tabName === 'saved') {
                renderSavedSpreadsheets();
            }
        });
    });
}

// ===========================
// CRIAR PLANILHAS
// ===========================

window.createSpreadsheet = function() {
    if (!currentUser) {
        showNotification('Você precisa estar logado!', 'error');
        return;
    }

    const name = document.getElementById('spreadsheet-name').value.trim();
    const description = document.getElementById('spreadsheet-description').value.trim();
    const numColumns = parseInt(document.getElementById('num-columns').value);
    const numRows = parseInt(document.getElementById('num-rows').value);

    if (!name) {
        showNotification('Por favor, informe um nome para a planilha', 'error');
        return;
    }

    if (numColumns < 1 || numColumns > 26) {
        showNotification('Número de colunas deve estar entre 1 e 26', 'error');
        return;
    }

    if (numRows < 1 || numRows > 1000) {
        showNotification('Número de linhas deve estar entre 1 e 1000', 'error');
        return;
    }

    const columns = [];
    for (let i = 0; i < numColumns; i++) {
        columns.push({
            name: `Coluna ${String.fromCharCode(65 + i)}`,
            type: DATA_TYPES.TEXT,
            required: false
        });
    }

    currentSpreadsheet = {
        id: Date.now().toString(),
        name: name,
        description: description,
        columns: columns,
        data: Array(numRows).fill(null).map(() => Array(numColumns).fill('')),
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
    };

    showEditor();
    showNotification('Planilha criada com sucesso!', 'success');
};

window.loadTemplate = function(templateKey) {
    if (!currentUser) {
        showNotification('Você precisa estar logado!', 'error');
        return;
    }

    const template = TEMPLATES[templateKey];
    if (!template) return;

    currentSpreadsheet = {
        id: Date.now().toString(),
        name: template.name,
        description: template.description,
        columns: JSON.parse(JSON.stringify(template.columns)),
        data: [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
    };

    for (let i = 0; i < template.rows; i++) {
        currentSpreadsheet.data.push(Array(template.columns.length).fill(''));
    }

    if (template.sampleData) {
        template.sampleData.forEach((rowData, index) => {
            if (index < currentSpreadsheet.data.length) {
                currentSpreadsheet.data[index] = [...rowData];
            }
        });
    }

    showEditor();
    showNotification(`Template "${template.name}" carregado!`, 'success');
};

// ===========================
// EDITOR
// ===========================

function showEditor() {
    document.getElementById('spreadsheet-editor').classList.remove('hidden');
    document.getElementById('current-spreadsheet-name').textContent = currentSpreadsheet.name;
    document.getElementById('current-spreadsheet-description').textContent = currentSpreadsheet.description;

    renderColumnSettings();
    renderTable();
    scrollToEditor();
}

function renderColumnSettings() {
    const container = document.getElementById('column-settings');
    container.innerHTML = '';

    currentSpreadsheet.columns.forEach((column, index) => {
        const columnDiv = document.createElement('div');
        columnDiv.className = 'column-setting';
        columnDiv.innerHTML = `
            <label>Coluna ${String.fromCharCode(65 + index)}</label>
            <input type="text"
                   value="${column.name}"
                   onchange="updateColumnName(${index}, this.value)"
                   placeholder="Nome da coluna">
            <select onchange="updateColumnType(${index}, this.value)">
                <option value="${DATA_TYPES.TEXT}" ${column.type === DATA_TYPES.TEXT ? 'selected' : ''}>Texto</option>
                <option value="${DATA_TYPES.NUMBER}" ${column.type === DATA_TYPES.NUMBER ? 'selected' : ''}>Número</option>
                <option value="${DATA_TYPES.DATE}" ${column.type === DATA_TYPES.DATE ? 'selected' : ''}>Data</option>
                <option value="${DATA_TYPES.EMAIL}" ${column.type === DATA_TYPES.EMAIL ? 'selected' : ''}>Email</option>
                <option value="${DATA_TYPES.CURRENCY}" ${column.type === DATA_TYPES.CURRENCY ? 'selected' : ''}>Moeda</option>
                <option value="${DATA_TYPES.PERCENTAGE}" ${column.type === DATA_TYPES.PERCENTAGE ? 'selected' : ''}>Porcentagem</option>
                <option value="${DATA_TYPES.FORMULA}" ${column.type === DATA_TYPES.FORMULA ? 'selected' : ''}>Fórmula</option>
            </select>
            <label style="margin-top: 8px; display: flex; align-items: center; gap: 5px;">
                <input type="checkbox"
                       ${column.required ? 'checked' : ''}
                       onchange="updateColumnRequired(${index}, this.checked)">
                Campo obrigatório
            </label>
        `;
        container.appendChild(columnDiv);
    });
}

function renderTable() {
    const table = document.getElementById('data-table');
    table.innerHTML = '';

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headerRow.innerHTML = '<th>#</th>';

    currentSpreadsheet.columns.forEach((column, index) => {
        const th = document.createElement('th');
        th.textContent = `${String.fromCharCode(65 + index)}: ${column.name}`;
        if (column.required) {
            th.innerHTML += ' <span style="color: #ff6b6b;">*</span>';
        }
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    currentSpreadsheet.data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');

        const tdNum = document.createElement('td');
        tdNum.textContent = rowIndex + 1;
        tdNum.style.background = '#f0f0f0';
        tdNum.style.fontWeight = 'bold';
        tr.appendChild(tdNum);

        row.forEach((cellValue, colIndex) => {
            const td = document.createElement('td');
            const column = currentSpreadsheet.columns[colIndex];

            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'cell-input';
            input.value = cellValue;
            input.dataset.row = rowIndex;
            input.dataset.col = colIndex;

            input.placeholder = getPlaceholderForType(column.type);

            input.addEventListener('input', function() {
                validateCell(this, rowIndex, colIndex);
            });

            input.addEventListener('blur', function() {
                updateCell(rowIndex, colIndex, this.value);
            });

            if (column.type === DATA_TYPES.FORMULA && column.formula) {
                input.disabled = true;
                input.style.background = '#f0f4ff';
                input.value = calculateFormula(column.formula, rowIndex);
            }

            td.appendChild(input);
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    recalculateAllFormulas();
}

function getPlaceholderForType(type) {
    const placeholders = {
        [DATA_TYPES.TEXT]: 'Digite o texto...',
        [DATA_TYPES.NUMBER]: 'Digite um número...',
        [DATA_TYPES.DATE]: 'DD/MM/AAAA ou AAAA-MM-DD',
        [DATA_TYPES.EMAIL]: 'email@exemplo.com',
        [DATA_TYPES.CURRENCY]: '1234.56',
        [DATA_TYPES.PERCENTAGE]: '50',
        [DATA_TYPES.FORMULA]: 'Calculado automaticamente'
    };
    return placeholders[type] || '';
}

// ===========================
// VALIDAÇÃO
// ===========================

function validateCell(input, rowIndex, colIndex) {
    const column = currentSpreadsheet.columns[colIndex];
    const value = input.value.trim();

    input.classList.remove('cell-error', 'cell-warning');
    clearCellErrors(rowIndex, colIndex);

    if (column.required && !value) {
        input.classList.add('cell-warning');
        showCellError(rowIndex, colIndex, 'Campo obrigatório', 'warning');
        return false;
    }

    if (!value) return true;

    let isValid = true;
    let errorMessage = '';

    switch (column.type) {
        case DATA_TYPES.NUMBER:
            if (isNaN(value)) {
                isValid = false;
                errorMessage = 'Deve ser um número válido';
            }
            break;

        case DATA_TYPES.DATE:
            if (!isValidDate(value)) {
                isValid = false;
                errorMessage = 'Data inválida. Use DD/MM/AAAA ou AAAA-MM-DD';
            }
            break;

        case DATA_TYPES.EMAIL:
            if (!isValidEmail(value)) {
                isValid = false;
                errorMessage = 'Email inválido';
            }
            break;

        case DATA_TYPES.CURRENCY:
            if (isNaN(value.replace(',', '.'))) {
                isValid = false;
                errorMessage = 'Valor monetário inválido';
            }
            break;

        case DATA_TYPES.PERCENTAGE:
            const num = parseFloat(value);
            if (isNaN(num) || num < 0 || num > 100) {
                isValid = false;
                errorMessage = 'Porcentagem deve estar entre 0 e 100';
            }
            break;
    }

    if (!isValid) {
        input.classList.add('cell-error');
        showCellError(rowIndex, colIndex, errorMessage, 'error');
    }

    return isValid;
}

function isValidDate(dateString) {
    const patterns = [
        /^\d{2}\/\d{2}\/\d{4}$/,
        /^\d{4}-\d{2}-\d{2}$/
    ];

    if (!patterns.some(p => p.test(dateString))) return false;

    const date = new Date(dateString);
    return date instanceof Date && !isNaN(date);
}

function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function showCellError(row, col, message, type = 'error') {
    const helpMessages = document.getElementById('help-messages');
    const errorId = `error-${row}-${col}`;

    const existing = document.getElementById(errorId);
    if (existing) existing.remove();

    const div = document.createElement('div');
    div.id = errorId;
    div.className = `help-message ${type}`;
    div.innerHTML = `<strong>${String.fromCharCode(65 + col)}${row + 1}:</strong> ${message}`;

    helpMessages.appendChild(div);
}

function clearCellErrors(row, col) {
    const errorId = `error-${row}-${col}`;
    const existing = document.getElementById(errorId);
    if (existing) existing.remove();
}

function updateCell(rowIndex, colIndex, value) {
    currentSpreadsheet.data[rowIndex][colIndex] = value;
    currentSpreadsheet.updatedAt = new Date().toISOString();
    recalculateAllFormulas();
}

// ===========================
// FÓRMULAS
// ===========================

function calculateFormula(formula, rowIndex) {
    try {
        let expression = formula.replace(/^=/, '');

        expression = expression.replace(/([A-Z])/g, (match, letter) => {
            const colIndex = letter.charCodeAt(0) - 65;
            if (colIndex < currentSpreadsheet.columns.length) {
                const value = currentSpreadsheet.data[rowIndex][colIndex];
                return value || '0';
            }
            return '0';
        });

        expression = expression.replace(/SOMA\(([^)]+)\)/gi, (match, range) => {
            return evaluateSum(range, rowIndex);
        });

        expression = expression.replace(/MEDIA\(([^)]+)\)/gi, (match, range) => {
            return evaluateAverage(range, rowIndex);
        });

        expression = expression.replace(/MAXIMO\(([^)]+)\)/gi, (match, range) => {
            return evaluateMax(range, rowIndex);
        });

        expression = expression.replace(/MINIMO\(([^)]+)\)/gi, (match, range) => {
            return evaluateMin(range, rowIndex);
        });

        expression = expression.replace(/CONTAR\(([^)]+)\)/gi, (match, range) => {
            return evaluateCount(range, rowIndex);
        });

        const result = eval(expression);
        return isNaN(result) ? '#ERRO!' : result.toString();

    } catch (error) {
        return '#ERRO!';
    }
}

function evaluateSum(range, rowIndex) {
    const values = getRangeValues(range, rowIndex);
    return values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
}

function evaluateAverage(range, rowIndex) {
    const values = getRangeValues(range, rowIndex);
    const sum = values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
    return values.length > 0 ? sum / values.length : 0;
}

function evaluateMax(range, rowIndex) {
    const values = getRangeValues(range, rowIndex);
    return Math.max(...values.map(v => parseFloat(v) || 0));
}

function evaluateMin(range, rowIndex) {
    const values = getRangeValues(range, rowIndex);
    return Math.min(...values.map(v => parseFloat(v) || 0));
}

function evaluateCount(range, rowIndex) {
    const values = getRangeValues(range, rowIndex);
    return values.filter(v => v !== '').length;
}

function getRangeValues(range, rowIndex) {
    const values = [];

    if (range.includes(':')) {
        const [start, end] = range.split(':');
        const startCol = start.charCodeAt(0) - 65;
        const endCol = end.charCodeAt(0) - 65;

        for (let col = startCol; col <= endCol; col++) {
            if (col < currentSpreadsheet.columns.length) {
                values.push(currentSpreadsheet.data[rowIndex][col]);
            }
        }
    } else {
        const colIndex = range.charCodeAt(0) - 65;
        if (colIndex < currentSpreadsheet.columns.length) {
            values.push(currentSpreadsheet.data[rowIndex][colIndex]);
        }
    }

    return values;
}

function recalculateAllFormulas() {
    currentSpreadsheet.columns.forEach((column, colIndex) => {
        if (column.type === DATA_TYPES.FORMULA && column.formula) {
            currentSpreadsheet.data.forEach((row, rowIndex) => {
                const result = calculateFormula(column.formula, rowIndex);
                const input = document.querySelector(`input[data-row="${rowIndex}"][data-col="${colIndex}"]`);
                if (input) {
                    input.value = result;
                }
            });
        }
    });
}

// ===========================
// ATUALIZAR COLUNAS
// ===========================

window.updateColumnName = function(index, name) {
    currentSpreadsheet.columns[index].name = name;
    currentSpreadsheet.updatedAt = new Date().toISOString();
    renderTable();
};

window.updateColumnType = function(index, type) {
    currentSpreadsheet.columns[index].type = type;
    currentSpreadsheet.updatedAt = new Date().toISOString();
    renderTable();
};

window.updateColumnRequired = function(index, required) {
    currentSpreadsheet.columns[index].required = required;
    currentSpreadsheet.updatedAt = new Date().toISOString();
    renderTable();
};

window.addRow = function() {
    const newRow = Array(currentSpreadsheet.columns.length).fill('');
    currentSpreadsheet.data.push(newRow);
    currentSpreadsheet.updatedAt = new Date().toISOString();
    renderTable();
    showNotification('Linha adicionada!', 'success');
};

window.addColumn = function() {
    if (currentSpreadsheet.columns.length >= 26) {
        showNotification('Máximo de 26 colunas atingido', 'error');
        return;
    }

    const newColumnIndex = currentSpreadsheet.columns.length;
    currentSpreadsheet.columns.push({
        name: `Coluna ${String.fromCharCode(65 + newColumnIndex)}`,
        type: DATA_TYPES.TEXT,
        required: false
    });

    currentSpreadsheet.data.forEach(row => row.push(''));
    currentSpreadsheet.updatedAt = new Date().toISOString();

    renderColumnSettings();
    renderTable();
    showNotification('Coluna adicionada!', 'success');
};

// ===========================
// SALVAR E CARREGAR
// ===========================

window.saveSpreadsheet = async function() {
    if (!currentSpreadsheet || !currentUser) return;

    currentSpreadsheet.updatedAt = new Date().toISOString();

    try {
        const firestoreId = await saveSpreadsheetToFirestore(currentSpreadsheet);

        if (firestoreId) {
            currentSpreadsheet.firestoreId = firestoreId;

            if (currentSpreadsheet.firestoreId) {
                await saveVersion(
                    currentSpreadsheet.firestoreId,
                    'Planilha salva',
                    JSON.parse(JSON.stringify(currentSpreadsheet))
                );
            }

            await loadUserSpreadsheets();
            renderSavedSpreadsheets();
        }
    } catch (error) {
        console.error('Erro ao salvar:', error);
    }
};

window.exportToExcel = function() {
    if (!currentSpreadsheet) return;

    const ws_data = [];
    const headers = currentSpreadsheet.columns.map(col => col.name);
    ws_data.push(headers);

    currentSpreadsheet.data.forEach(row => {
        ws_data.push([...row]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    XLSX.utils.book_append_sheet(wb, ws, "Dados");

    const fileName = `${currentSpreadsheet.name.replace(/[^a-z0-9]/gi, '_')}.xlsx`;
    XLSX.writeFile(wb, fileName);

    showNotification('Planilha exportada com sucesso!', 'success');
};

function renderSavedSpreadsheets() {
    const container = document.getElementById('saved-spreadsheets-list');
    container.innerHTML = '';

    if (spreadsheets.length === 0) {
        container.innerHTML = '<p class="help-text">Nenhuma planilha salva ainda. Crie sua primeira planilha!</p>';
        return;
    }

    spreadsheets.forEach(spreadsheet => {
        const div = document.createElement('div');
        div.className = 'spreadsheet-item';

        const updatedDate = spreadsheet.updatedAt?.toDate ?
            spreadsheet.updatedAt.toDate().toLocaleString('pt-BR') :
            new Date(spreadsheet.updatedAt).toLocaleString('pt-BR');

        const isOwner = spreadsheet.owner === currentUser.uid;
        const permission = spreadsheet.permissions?.[currentUser.uid] || 'viewer';

        div.innerHTML = `
            <div class="spreadsheet-info">
                <h3>${spreadsheet.name}</h3>
                <p>${spreadsheet.description || 'Sem descrição'}</p>
                <div class="spreadsheet-meta">
                    ${spreadsheet.columns.length} colunas × ${spreadsheet.data.length} linhas |
                    Atualizado em ${updatedDate}
                    ${!isOwner ? `| <span class="collaborator-permission permission-${permission}">${permission}</span>` : ''}
                </div>
            </div>
            <div class="spreadsheet-actions">
                <button class="btn btn-small btn-primary" onclick="openSpreadsheet('${spreadsheet.firestoreId}')">Abrir</button>
                ${isOwner ? `
                    <button class="btn btn-small btn-secondary" onclick="duplicateSpreadsheet('${spreadsheet.firestoreId}')">Duplicar</button>
                    <button class="btn btn-small" style="background: #f44336; color: white;" onclick="deleteSpreadsheet('${spreadsheet.firestoreId}')">Excluir</button>
                ` : ''}
            </div>
        `;

        container.appendChild(div);
    });
}

window.openSpreadsheet = async function(firestoreId) {
    const spreadsheet = await loadSpreadsheetFromFirestore(firestoreId);
    if (!spreadsheet) return;

    currentSpreadsheet = JSON.parse(JSON.stringify(spreadsheet));
    versionHistory = await loadVersionHistory(firestoreId);

    showEditor();
    showNotification(`Planilha "${spreadsheet.name}" aberta!`, 'success');
};

window.duplicateSpreadsheet = async function(firestoreId) {
    const spreadsheet = await loadSpreadsheetFromFirestore(firestoreId);
    if (!spreadsheet) return;

    const duplicate = JSON.parse(JSON.stringify(spreadsheet));
    duplicate.id = Date.now().toString();
    duplicate.name = `${spreadsheet.name} (Cópia)`;
    duplicate.createdAt = new Date().toISOString();
    duplicate.updatedAt = new Date().toISOString();
    delete duplicate.firestoreId;

    currentSpreadsheet = duplicate;

    await saveSpreadsheetToFirestore(duplicate);
    await loadUserSpreadsheets();
    renderSavedSpreadsheets();

    showNotification('Planilha duplicada com sucesso!', 'success');
};

window.deleteSpreadsheet = async function(firestoreId) {
    if (!confirm('Tem certeza que deseja excluir esta planilha?')) return;

    const success = await deleteSpreadsheetFromFirestore(firestoreId);
    if (success) {
        await loadUserSpreadsheets();
        renderSavedSpreadsheets();
    }
};

window.closeEditor = function() {
    if (confirm('Deseja salvar antes de fechar?')) {
        saveSpreadsheet();
    }

    document.getElementById('spreadsheet-editor').classList.add('hidden');
    currentSpreadsheet = null;
};

// ===========================
// COMPARTILHAMENTO
// ===========================

window.showShareModal = async function() {
    if (!currentSpreadsheet || !currentSpreadsheet.firestoreId) {
        showNotification('Salve a planilha antes de compartilhar', 'warning');
        return;
    }

    const permission = currentSpreadsheet.permissions?.[currentUser.uid];
    if (permission !== 'owner' && permission !== 'editor') {
        showNotification('Apenas o dono ou editores podem compartilhar', 'error');
        return;
    }

    document.getElementById('share-modal').classList.remove('hidden');
    await loadCollaborators();
};

window.closeShareModal = function() {
    document.getElementById('share-modal').classList.add('hidden');
};

async function loadCollaborators() {
    if (!currentSpreadsheet || !currentSpreadsheet.firestoreId) return;

    const collaborators = await getSpreadsheetCollaborators(currentSpreadsheet.firestoreId);
    const container = document.getElementById('collaborators-list');
    container.innerHTML = '';

    if (collaborators.length === 0) {
        container.innerHTML = '<p class="help-text">Nenhum colaborador ainda</p>';
        return;
    }

    collaborators.forEach(collab => {
        const div = document.createElement('div');
        div.className = 'collaborator-item';

        const canRemove = currentSpreadsheet.owner === currentUser.uid && collab.uid !== currentUser.uid;

        div.innerHTML = `
            <div class="collaborator-info">
                <img src="${collab.photoURL || 'https://via.placeholder.com/40'}" alt="${collab.displayName}" class="collaborator-photo">
                <div class="collaborator-details">
                    <div class="collaborator-name">${collab.displayName}</div>
                    <div class="collaborator-email">${collab.email}</div>
                </div>
                <span class="collaborator-permission permission-${collab.permission}">${collab.permission}</span>
            </div>
            ${canRemove ? `<button class="btn-remove" onclick="handleRemoveAccess('${collab.uid}')">Remover</button>` : ''}
        `;

        container.appendChild(div);
    });
}

window.handleShareSpreadsheet = async function() {
    const email = document.getElementById('share-email').value.trim();
    const permission = document.getElementById('share-permission').value;

    if (!email) {
        showNotification('Digite um email', 'error');
        return;
    }

    if (!currentSpreadsheet || !currentSpreadsheet.firestoreId) {
        showNotification('Salve a planilha primeiro', 'error');
        return;
    }

    const success = await shareSpreadsheet(currentSpreadsheet.firestoreId, email, permission);

    if (success) {
        document.getElementById('share-email').value = '';
        await loadCollaborators();
    }
};

window.handleRemoveAccess = async function(userId) {
    if (!confirm('Deseja remover o acesso deste usuário?')) return;

    const success = await removeUserAccess(currentSpreadsheet.firestoreId, userId);

    if (success) {
        await loadCollaborators();
    }
};

// ===========================
// HISTÓRICO DE VERSÕES
// ===========================

window.showVersionHistory = async function() {
    const modal = document.getElementById('version-modal');
    const versionList = document.getElementById('version-list');

    versionList.innerHTML = '<p class="help-text">Carregando histórico...</p>';
    modal.classList.remove('hidden');

    if (currentSpreadsheet && currentSpreadsheet.firestoreId) {
        const versions = await loadVersionHistory(currentSpreadsheet.firestoreId);

        versionList.innerHTML = '';

        if (versions.length === 0) {
            versionList.innerHTML = '<p class="help-text">Nenhuma versão no histórico</p>';
        } else {
            versions.forEach((version) => {
                const div = document.createElement('div');
                div.className = 'version-item';

                const date = version.timestamp?.toDate ?
                    version.timestamp.toDate().toLocaleString('pt-BR') :
                    new Date(version.timestamp).toLocaleString('pt-BR');

                div.innerHTML = `
                    <div class="version-date">${date}</div>
                    <div class="version-changes">${version.action} - por ${version.userName}</div>
                `;

                versionList.appendChild(div);
            });
        }
    }
};

window.closeVersionModal = function() {
    document.getElementById('version-modal').classList.add('hidden');
};

// ===========================
// IMPORTAÇÃO
// ===========================

function initializeUploadArea() {
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');

    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileUpload(files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileUpload(e.target.files[0]);
        }
    });
}

function handleFileUpload(file) {
    if (!currentUser) {
        showNotification('Você precisa estar logado!', 'error');
        return;
    }

    const fileName = file.name.toLowerCase();

    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
        showNotification('Por favor, selecione um arquivo Excel (.xlsx, .xls) ou CSV', 'error');
        return;
    }

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            if (jsonData.length === 0) {
                showNotification('Arquivo vazio ou inválido', 'error');
                return;
            }

            const headers = jsonData[0];
            const dataRows = jsonData.slice(1);

            currentSpreadsheet = {
                id: Date.now().toString(),
                name: file.name.replace(/\.[^/.]+$/, ""),
                description: `Importado de ${file.name}`,
                columns: headers.map(header => ({
                    name: header || 'Sem nome',
                    type: DATA_TYPES.TEXT,
                    required: false
                })),
                data: dataRows.map(row => {
                    const paddedRow = [...row];
                    while (paddedRow.length < headers.length) {
                        paddedRow.push('');
                    }
                    return paddedRow.map(cell => cell !== undefined && cell !== null ? cell.toString() : '');
                }),
                createdAt: new Date().toISOString(),
                updatedAt: new Date().toISOString()
            };

            showEditor();
            showNotification('Arquivo importado com sucesso!', 'success');

        } catch (error) {
            showNotification('Erro ao importar arquivo: ' + error.message, 'error');
        }
    };

    reader.readAsArrayBuffer(file);
}

// ===========================
// UTILITÁRIOS
// ===========================

function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `help-message ${type}`;
    notification.textContent = message;
    notification.style.position = 'fixed';
    notification.style.top = '20px';
    notification.style.right = '20px';
    notification.style.zIndex = '10000';
    notification.style.minWidth = '300px';
    notification.style.boxShadow = '0 4px 12px rgba(0,0,0,0.3)';

    document.body.appendChild(notification);

    setTimeout(() => {
        notification.style.opacity = '0';
        notification.style.transition = 'opacity 0.3s';
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

function scrollToEditor() {
    document.getElementById('spreadsheet-editor').scrollIntoView({
        behavior: 'smooth',
        block: 'start'
    });
}

window.showNotification = showNotification;

console.log('Sistema de Gestão de Planilhas carregado com sucesso!');
