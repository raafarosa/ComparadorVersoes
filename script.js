let filesState = { hml: [], preprod: [], prod: [] };
let resultados = { hml: null, preprod: null, prod: null };
let currentMode = 'auto';
let tableDataCache = {
    preProd: [],
    preHml: [],
    hmlProd: [],
    duplicadas: []
};

// Dicionário de sinônimos/aliases trazido do script do seu amigo para mapear planilhas dinamicamente
const aliases = {
    appName: ["aplicacao", "aplicação", "application", "nome", "nomeaplicacao", "nome da aplicacao", "nome da aplicação", "app", "sistema"],
    changeset: ["changeset", "change set", "change_set", "change"],
    release: ["release", "versao", "versão", "version"],
};

const autoDrop = document.getElementById('autoDrop');
const autoFileInput = document.getElementById('autoFileInput');
const autoFileList = document.getElementById('autoFileList');
const autoStatus = document.getElementById('autoStatus');

// Funções auxiliares de normalização e busca por apelidos (XLSX)
function normalizeLabel(value) {
    return String(value || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-zA-Z0-9]/g, "")
        .toLowerCase();
}

function findValueByAlias(obj, kind) {
    const keys = Object.keys(obj || {});
    const wanted = aliases[kind].map(normalizeLabel);
    for (const key of keys) {
        const normalized = normalizeLabel(key);
        if (wanted.includes(normalized)) {
            return obj[key];
        }
    }
    return undefined;
}

// Converte as linhas extraídas do XLSX para a estrutura padronizada de Aplicações do script
function parseXlsxRows(rows) {
    const aplicacoes = [];
    for (const row of rows) {
        const appName = findValueByAlias(row, "appName");
        if (!appName) continue;

        const changeset = findValueByAlias(row, "changeset") ?? "0";
        const release = findValueByAlias(row, "release") ?? "0";

        // Planilhas costumam não ter a coluna Path física, simulamos uma chave para compatibilidade
        const simuladoPath = `\\\\XLSX_GENERATED\\\\${String(appName).trim()}`;

        aplicacoes.push({
            Path: simuladoPath,
            Changeset: String(changeset).trim(),
            Release: String(release).trim()
        });
    }
    return { Aplicacoes: aplicacoes };
}

// Processa o buffer do Excel lendo todas as abas existentes
function parseWorkbook(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const allRows = [];
    wb.SheetNames.forEach((name) => {
        const sheet = wb.Sheets[name];
        // blankrows: false impede que linhas fantasmas quebrem o mapeador
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", blankrows: false });
        if (rows.length > 0) {
            allRows.push(...rows);
        }
    });
    return allRows;
}

// Mantém a leitura do formato texto/TSV do SaaS limpo
function parseSaasTextFile(text) {
    const lines = text.split('\n');
    const aplicacoes = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line || line.startsWith('Aplicação') || line.toLowerCase().includes('changeset')) {
            continue;
        }

        const parts = line.split(/\t| {2,}/);
        if (parts.length > 0) {
            const aplicacaoCrua = parts[0].trim();
            const changeset = parts[1] ? parts[1].trim() : '0';
            const release = parts[2] ? parts[2].trim() : '';
            const simuladoPath = `\\\\SAAS_GENERATED\\\\${aplicacaoCrua}`;

            aplicacoes.push({
                Path: simuladoPath,
                Changeset: changeset,
                Release: release
            });
        }
    }
    return { Aplicacoes: aplicacoes };
}

function extractAppsFromJson(json, servidorPadrao) {
    const apps = [];
    const aplicacoes = json.Aplicacoes || json.aplicacoes || json.Aplicações || [];
    const servidor = json.Servidor || json.servidor || servidorPadrao;

    for (const app of aplicacoes) {
        if (!app.Path) continue;

        const upperPath = app.Path.toUpperCase();
        
        // --- FILTRO DE DIRETÓRIOS E PASTAS INDEVIDAS ---
        if (
            upperPath.includes('VERSIONAMENTO PARALELO HML') || 
            upperPath.includes('\\VERSIONAMENTO PARALELO\\') ||
            upperPath.includes('\\COMPATIBILIDADECNPJALFANUMERICO') || // Captura o termo em qualquer parte do path
            upperPath.includes('\\FI.DEPLOY') || // Captura FI.Deploy com ou sem contrabarra no final
            upperPath.includes('\\BKP\\') || upperPath.endsWith('\\BKP') ||
            upperPath.includes('\\OLD\\') || upperPath.endsWith('\\OLD') ||
            upperPath.includes('\\BACKUP\\') || upperPath.endsWith('\\BACKUP') ||
            upperPath.includes('\\SUBIDA\\') ||
            upperPath.includes('\\_2\\') || upperPath.endsWith('\\_2')
        ) {
            continue; 
        }

        let parts = app.Path.split('\\').filter(p => p.trim() !== '');
        if (parts.length === 0) continue;

        let name = parts[parts.length - 1];

        // --- TRAVA DE NOME DE APLICAÇÃO INDEVIDA / DUPLICADA ---
        // 1. Remove variações do validador de CNPJ Alfanumérico antes de tratar o nome
        if (name.toUpperCase().startsWith('FI.VERIFICACOMPATIBILIDADECNPJALFANUMERICO')) {
            continue;
        }

        // 2. Se o nome terminar com underline seguido de qualquer número (ex: WebBackofficeDC_2, App_3), ignora.
        if (/_\d+(\.exe)?$/i.test(name)) {
            continue; 
        }

        // Remove ID numérico inicial padrão SaaS/Planilhas (Ex: "28421 - Web CDC Contábil" vira "Web CDC Contábil")
        name = name.replace(/^\d+\s*-\s*/, '');

        name = name.replace(/_I\d+_R_C\d+/gi, '');
        name = name.replace(/_C\d+$/gi, '');
        name = name.replace(/\.exe$/i, '');
        name = name.replace(/_\d+$/i, '');
        name = name.replace(/_old$/i, '').replace(/_bkp$/i, '').replace(/_backup$/i);
        name = name.replace(/_2$/i, '').replace(/_3$/i, '').replace(/_4$/i);

        const norm = {
            'webbacenifrs': 'WebBacen_IFRS',
            'webcorporatecontabilifrs': 'WebCorporateContabil_IFRS',
            'webtabelasifrs': 'WebTabelas_IFRS',
            'webbackofficedc': 'WebBackofficeDC',
            'webcob': 'WebCCob',
            'webgedoc': 'WebGEDoc',
            'wsautcorporate': 'WsAutCorporate',
            'wscnab': 'WsCNAB',
            'wsduplocontrole': 'WSDuploControle',
            'wsged': 'WSGED',
            'wsmonitor': 'WsMonitor',
            'wssic': 'WsSIC'
        };
        if (norm[name.toLowerCase()]) name = norm[name.toLowerCase()];

        const changeset = app.Changeset || app.changeset || '0';
        const release = app.Release || app.release || '';
        let versao = '';
        if (release && release.trim() !== '' && release !== '0') {
            versao = `C${changeset} | R${release}`;
        } else if (changeset && changeset !== '0') {
            versao = `C${changeset}`;
        } else {
            versao = 'Sem versão';
        }

        apps.push({
            nome: name,
            path: app.Path,
            versao: versao,
            changeset: changeset,
            release: release,
            servidor: servidor
        });
    }
    return apps;
}

async function processEnvironment(envName, files) {
    const appsMap = new Map();
    const duplicadas = [];

    for (const file of files) {
        try {
            const ext = file.name.split('.').pop().toLowerCase();
            let json;

            if (ext === "json") {
                const text = await file.text();
                const trimmedText = text.trim();
                if (trimmedText.startsWith('{') || trimmedText.startsWith('[')) {
                    json = JSON.parse(text);
                } else {
                    json = parseSaasTextFile(text);
                }
            } else if (["xlsx", "xls"].includes(ext)) {
                const arrayBuffer = await file.arrayBuffer();
                const rawRows = parseWorkbook(arrayBuffer);
                json = parseXlsxRows(rawRows);
            } else {
                const text = await file.text();
                json = parseSaasTextFile(text);
            }

            const apps = extractAppsFromJson(json, envName);

            for (const app of apps) {
                const key = `${app.nome}|${app.path}`;

                if (appsMap.has(key)) {
                    const existing = appsMap.get(key);
                    duplicadas.push({
                        aplicacao: app.nome,
                        path: app.path,
                        versao1: existing.versao,
                        versao2: app.versao,
                        ambiente: envName
                    });
                    const num1 = parseInt(existing.changeset) || 0;
                    const num2 = parseInt(app.changeset) || 0;
                    if (num2 > num1) {
                        appsMap.set(key, app);
                    }
                } else {
                    appsMap.set(key, app);
                }
            }
        } catch (e) {
            console.error(`Erro ao processar ${file.name}:`, e);
        }
    }

    const appsByName = new Map();
    for (const [_, app] of appsMap) {
        if (!appsByName.has(app.nome)) {
            appsByName.set(app.nome, []);
        }
        appsByName.get(app.nome).push(app);
    }
    return { apps: appsByName, duplicadas };
}

function avaliarFormulaExcel(b3, c3, labelAmbienteA, labelAmbienteB) {
    const b = (b3 || "").toString().trim().toUpperCase();
    const c = (c3 || "").toString().trim().toUpperCase();

    if (b === "NÃO ENCONTRADO" || c === "NÃO ENCONTRADO" || b === "N/A" || c === "N/A" || b === "" || c === "") {
        return "N/D";
    }

    const extrairPrincipal = (texto) => {
        if (texto.length < 2) return 0;
        const textoCortado = texto.substring(1);
        const espacoIdx = textoCortado.indexOf(" ");
        const parteNumerica = espacoIdx === -1 ? textoCortado : textoCortado.substring(0, espacoIdx);
        const valor = parseInt(parteNumerica, 10);
        return isNaN(valor) ? 0 : valor;
    };

    const extrairRevisao = (texto) => {
        const rIdx = texto.indexOf("R");
        if (rIdx === -1) return 0;
        const parteNumerica = texto.substring(rIdx + 1);
        const valor = parseInt(parteNumerica, 10);
        return isNaN(valor) ? 0 : valor;
    };

    const numPrePrincipal = extrairPrincipal(b);
    const numProdPrincipal = extrairPrincipal(c);

    if (numPrePrincipal > numProdPrincipal) return `Maior em ${labelAmbienteA}`;
    if (numPrePrincipal < numProdPrincipal) return `Maior em ${labelAmbienteB}`;

    const numPreRevisao = extrairRevisao(b);
    const numProdRevisao = extrairRevisao(c);

    if (numPreRevisao > numProdRevisao) return `Maior em ${labelAmbienteA}`;
    if (numPreRevisao < numProdRevisao) return `Maior em ${labelAmbienteB}`;

    return "Versões Iguais";
}

function compareEnvironments(envA, envB, nameA, nameB) {
    const results = [];
    const allNames = new Set();

    for (const nome of envA.keys()) allNames.add(nome);
    for (const nome of envB.keys()) allNames.add(nome);

    for (const nome of allNames) {
        const appsA = envA.get(nome) || [];
        const appsB = envB.get(nome) || [];

        if (appsA.length > 1 || appsB.length > 1) {
            for (const appA of appsA) {
                let matched = false;
                for (const appB of appsB) {
                    if (appA.path === appB.path) {
                        matched = true;
                        let status = avaliarFormulaExcel(appA.versao, appB.versao, nameA, nameB);
                        results.push({
                            nome: nome,
                            versaoA: appA.versao,
                            versaoB: appB.versao,
                            status: status,
                            pathA: appA.path,
                            pathB: appB.path
                        });
                        break;
                    }
                }
                if (!matched) {
                    results.push({
                        nome: nome,
                        versaoA: appA.versao,
                        versaoB: 'N/A',
                        status: `Só em ${nameA}`,
                        pathA: appA.path,
                        pathB: 'N/D'
                    });
                }
            }
            for (const appB of appsB) {
                let found = false;
                for (const appA of appsA) {
                    if (appA.path === appB.path) found = true;
                }
                if (!found) {
                    results.push({
                        nome: nome,
                        versaoA: 'N/A',
                        versaoB: appB.versao,
                        status: `Só em ${nameB}`,
                        pathA: 'N/D',
                        pathB: appB.path
                    });
                }
            }
        } else if (appsA.length === 1 && appsB.length === 1) {
            const appA = appsA[0];
            const appB = appsB[0];
            let status = avaliarFormulaExcel(appA.versao, appB.versao, nameA, nameB);
            results.push({
                nome: nome,
                versaoA: appA.versao,
                versaoB: appB.versao,
                status: status,
                pathA: appA.path,
                pathB: appB.path
            });
        } else if (appsA.length === 0 && appsB.length > 0) {
            for (const appB of appsB) {
                results.push({
                    nome: nome,
                    versaoA: 'N/A',
                    versaoB: appB.versao,
                    status: `Só em ${nameB}`,
                    pathA: 'N/D',
                    pathB: appB.path
                });
            }
        } else if (appsA.length > 0 && appsB.length === 0) {
            for (const appA of appsA) {
                results.push({
                    nome: nome,
                    versaoA: appA.versao,
                    versaoB: 'N/A',
                    status: `Só em ${nameA}`,
                    pathA: appA.path,
                    pathB: 'N/D'
                });
            }
        }
    }

    results.sort((a, b) => a.nome.localeCompare(b.nome));
    return results;
}

function getStatusClass(status) {
    if (status === 'Versões Iguais') return 'diff-ok';
    if (status.startsWith('Maior em HML') || status.startsWith('Maior em PREPROD')) return 'diff-pre-maior';
    if (status.startsWith('Maior em PROD')) return 'diff-prod-maior';
    if (status.includes('Só em')) return 'diff-only';
    if (status === 'N/D') return 'diff-nd';
    return '';
}

function renderTableWithFilters(tableId, data, columns, tableType) {
    const table = document.getElementById(tableId);
    if (!table) return;

    tableDataCache[tableType] = data;

    const thead = table.querySelector('thead');
    thead.innerHTML = '';

    const headerRow = document.createElement('tr');
    columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        th.style.background = '#2a2a2a';
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    const filterRow = document.createElement('tr');
    filterRow.className = 'filter-row';
    columns.forEach((col, idx) => {
        const th = document.createElement('th');
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = `Filtrar ${col}...`;
        input.className = 'filter-input';
        input.dataset.column = idx;
        input.addEventListener('keyup', () => applyFilters());
        th.appendChild(input);
        filterRow.appendChild(th);
    });
    thead.appendChild(filterRow);

    const tbody = table.querySelector('tbody');

    function applyFilters() {
        const filters = [];
        const filterInputs = filterRow.querySelectorAll('.filter-input');
        filterInputs.forEach(input => {
            filters.push(input.value.toLowerCase().trim());
        });

        const filteredData = data.filter(row => {
            const rowValues = Object.values(row);
            for (let i = 0; i < filters.length; i++) {
                if (filters[i] && !String(rowValues[i] || '').toLowerCase().includes(filters[i])) {
                    return false;
                }
            }
            return true;
        });

        tbody.innerHTML = '';
        for (const row of filteredData) {
            const tr = document.createElement('tr');
            const rowValues = Object.values(row);
            rowValues.forEach((val, idx) => {
                const td = document.createElement('td');
                if (idx === 3 && columns.length > 3 && columns[3] === 'Status') {
                    td.className = getStatusClass(val);
                } else if (idx >= 4 && (val === 'N/D' || val === 'N/A')) {
                    td.style.color = '#888888';
                    td.style.fontStyle = 'italic';
                }
                td.textContent = val || '';
                if (idx >= 4) td.style.wordBreak = 'break-all';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        }
    }

    applyFilters();

    const clearBtn = document.querySelector(`.clear-filters-btn[data-table="${tableId}"]`);
    if (clearBtn) {
        clearBtn.onclick = () => {
            filterRow.querySelectorAll('.filter-input').forEach(input => input.value = '');
            applyFilters();
        };
    }
}

function updateCompareButton() {
    const hasHml = filesState.hml && filesState.hml.length > 0;
    const hasPre = filesState.preprod && filesState.preprod.length > 0;
    const hasProd = filesState.prod && filesState.prod.length > 0;

    const activeEnvs = [hasHml, hasPre, hasProd].filter(Boolean).length;
    
    const btnCompare = document.getElementById('btnCompare');
    if (btnCompare) {
        if (activeEnvs >= 2) {
            btnCompare.disabled = false;
            btnCompare.removeAttribute('disabled');
            btnCompare.style.background = '#1a73e8';
            btnCompare.style.cursor = 'pointer';
        } else {
            btnCompare.disabled = true;
            btnCompare.style.background = '';
            btnCompare.style.cursor = 'not-allowed';
        }
    }
}

function detectEnvByFilename(filename) {
    const upper = filename.toUpperCase();
    if (upper.includes('HML') || upper.includes('HOMOLOG')) return 'hml';
    if (upper.includes('PREP') || upper.includes('PREPROD') || upper.includes('PRE-PROD')) return 'preprod';
    if (upper.includes('PROD') || upper.includes('PRODUCAO') || upper.includes('PRD')) return 'prod';
    return null;
}

// ========================================================
// GERENCIAMENTO DE CAPTURA DE EVENTOS (AUTOMÁTICO E MANUAL)
// ========================================================

// 1. Evento unificado do Modo Automático
if (autoDrop && autoFileInput) {
    autoDrop.addEventListener('click', (e) => {
        if (e.target === autoFileInput) return;
        autoFileInput.click();
    });

    autoFileInput.addEventListener('change', (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        filesState = { hml: [], preprod: [], prod: [] };
        autoStatus.innerHTML = '';
        let identified = { hml: 0, preprod: 0, prod: 0, unknown: 0 };

        for (const file of files) {
            const env = detectEnvByFilename(file.name);
            if (env) {
                filesState[env].push(file);
                identified[env]++;
                autoStatus.innerHTML += `<span class="file-tag" style="border-color:#5adaaa; padding:2px 6px; margin:2px; display:inline-block; border:1px solid;">✅ ${file.name} → ${env.toUpperCase()}</span> `;
            } else {
                identified.unknown++;
                autoStatus.innerHTML += `<span class="file-tag" style="border-color:#ff8888; padding:2px 6px; margin:2px; display:inline-block; border:1px solid;">⚠️ ${file.name} → não identificado</span> `;
            }
        }

        autoFileList.innerHTML = `
            <br><strong>📊 Resumo dos Arquivos:</strong><br>
            🏗️ HML: ${identified.hml} arquivo(s)<br>
            ⚙️ PREPROD: ${identified.preprod} arquivo(s)<br>
            🚀 PROD: ${identified.prod} arquivo(s)<br>
            ${identified.unknown > 0 ? `⚠️ Não identificados: ${identified.unknown}` : ''}
        `;
        
        updateCompareButton();
    });
}

// 2. Evento unificado e corrigido do Modo Manual (Removido loops fantasmas)
const manualInputs = document.querySelectorAll('.manual-file-input');
manualInputs.forEach(input => {
    const env = input.dataset.env;
    const dropArea = input.closest('.file-drop');

    if (dropArea) {
        dropArea.addEventListener('click', (e) => {
            if (e.target === input) return;
            input.click();
        });
    }

    input.addEventListener('change', (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        filesState[env] = files;
        
        const listDiv = document.getElementById(`${env}FileList`);
        if (listDiv) {
            listDiv.innerHTML = files.map(f => `<span class="file-tag" style="padding:2px 6px; margin:2px; display:inline-block; border:1px solid #444;">📄 ${f.name}</span>`).join('');
        }
        
        updateCompareButton();
    });
});

// Selector de Abas / Modos de Painel
document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.mode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentMode = btn.dataset.mode;
        document.getElementById('auto-panel').classList.toggle('active', currentMode === 'auto');
        document.getElementById('manual-panel').classList.toggle('active', currentMode === 'manual');
    });
});

// Ação do Botão Comparador
document.getElementById('btnCompare').addEventListener('click', async () => {
    const loading = document.getElementById('loading');
    const progressFill = document.getElementById('progressFill');
    loading.style.display = 'block';
    progressFill.style.width = '0%';

    try {
        progressFill.style.width = '20%';
        const hmlResult = filesState.hml.length > 0 ? await processEnvironment('HML', filesState.hml) : { apps: new Map(), duplicadas: [] };
        progressFill.style.width = '50%';
        const preResult = filesState.preprod.length > 0 ? await processEnvironment('PREPROD', filesState.preprod) : { apps: new Map(), duplicadas: [] };
        progressFill.style.width = '80%';
        const prodResult = filesState.prod.length > 0 ? await processEnvironment('PROD', filesState.prod) : { apps: new Map(), duplicadas: [] };
        progressFill.style.width = '100%';

        resultados = { hml: hmlResult.apps, preprod: preResult.apps, prod: prodResult.apps };

        let preProdData = [], preHmlData = [], hmlProdData = [];
        let diffPreProd = 0, diffPreHml = 0;

        if (filesState.preprod.length > 0 && filesState.prod.length > 0) {
            const preProd = compareEnvironments(resultados.preprod, resultados.prod, 'PREPROD', 'PROD');
            diffPreProd = preProd.filter(r => r.status.startsWith('Maior em')).length;
            preProdData = preProd.map(r => ({
                Aplicação: r.nome,
                VersaoPREPROD: r.versaoA,
                VersaoPROD: r.versaoB,
                Status: r.status,
                PathPREPROD: r.pathA,
                PathPROD: r.pathB
            }));
        }

        if (filesState.preprod.length > 0 && filesState.hml.length > 0) {
            const preHml = compareEnvironments(resultados.preprod, resultados.hml, 'PREPROD', 'HML');
            diffPreHml = preHml.filter(r => r.status.startsWith('Maior em')).length;
            preHmlData = preHml.map(r => ({
                Aplicação: r.nome,
                VersaoPREPROD: r.versaoA,
                VersaoHML: r.versaoB,
                Status: r.status,
                PathPREPROD: r.pathA,
                PathHML: r.pathB
            }));
        }

        if (filesState.hml.length > 0 && filesState.prod.length > 0) {
            const hmlProd = compareEnvironments(resultados.hml, resultados.prod, 'HML', 'PROD');
            hmlProdData = hmlProd.map(r => ({
                Aplicação: r.nome,
                VersaoHML: r.versaoA,
                VersaoPROD: r.versaoB,
                Status: r.status,
                PathHML: r.pathA,
                PathPROD: r.pathB
            }));
        }

        renderTableWithFilters('tablePreProd', preProdData, ['Aplicação', 'Versão PREPROD', 'Versão PROD', 'Status', 'Path PREPROD', 'Path PROD'], 'preProd');
        renderTableWithFilters('tablePreHml', preHmlData, ['Aplicação', 'Versão PREPROD', 'Versão HML', 'Status', 'Path PREPROD', 'Path HML'], 'preHml');
        renderTableWithFilters('tableHmlProd', hmlProdData, ['Aplicação', 'Versão HML', 'Versão PROD', 'Status', 'Path HML', 'Path PROD'], 'hmlProd');

        const todasDups = [...hmlResult.duplicadas, ...preResult.duplicadas, ...prodResult.duplicadas];
        const duplicadasData = todasDups.map(d => ({
            Ambiente: d.ambiente || '?',
            Aplicação: d.aplicacao,
            Path: d.path,
            Versao1: d.versao1,
            Versao2: d.versao2
        }));
        renderTableWithFilters('tableDuplicadas', duplicadasData, ['Ambiente', 'Aplicação', 'Path', 'Versão 1', 'Versão 2'], 'duplicadas');

        const totalAppsHml = resultados.hml ? Array.from(resultados.hml.values()).reduce((acc, arr) => acc + arr.length, 0) : 0;
        const totalAppsPre = resultados.preprod ? Array.from(resultados.preprod.values()).reduce((acc, arr) => acc + arr.length, 0) : 0;
        const totalAppsProd = resultados.prod ? Array.from(resultados.prod.values()).reduce((acc, arr) => acc + arr.length, 0) : 0;

        document.getElementById('stats').innerHTML = `
            <div class="stat-card" style="display: ${totalAppsHml > 0 ? 'block' : 'none'}"><div class="stat-number" style="color:#7a9aba;">${totalAppsHml}</div><div class="stat-label">Apps HML</div></div>
            <div class="stat-card" style="display: ${totalAppsPre > 0 ? 'block' : 'none'}"><div class="stat-number" style="color:#daaa5a;">${totalAppsPre}</div><div class="stat-label">Apps PREPROD</div></div>
            <div class="stat-card" style="display: ${totalAppsProd > 0 ? 'block' : 'none'}"><div class="stat-number" style="color:#5adaaa;">${totalAppsProd}</div><div class="stat-label">Apps PROD</div></div>
            <div class="stat-card" style="display: ${preProdData.length > 0 ? 'block' : 'none'}"><div class="stat-number" style="color:#ffaaaa;">${diffPreProd}</div><div class="stat-label">Divergências PRE/PROD</div></div>
            <div class="stat-card" style="display: ${preHmlData.length > 0 ? 'block' : 'none'}"><div class="stat-number" style="color:#ffaaaa;">${diffPreHml}</div><div class="stat-label">Divergências PRE/HML</div></div>
        `;

        document.getElementById('resultSection').style.display = 'block';
        document.getElementById('btnExport').disabled = false;
    } catch (err) {
        alert('Erro: ' + err.message);
        console.error(err);
    } finally {
        loading.style.display = 'none';
    }
});

// Função de Exportação CSV
document.getElementById('btnExport').addEventListener('click', () => {
    const preProdData = tableDataCache.preProd || [];
    const preHmlData = tableDataCache.preHml || [];
    const hmlProdData = tableDataCache.hmlProd || [];

    function arrayToCSV(data, headers) {
        const rows = [headers.join(',')];

        const normalizeKey = (header) => {
            return header
                .normalize("NFD")
                .replace(/[\u0300-\u036f]/g, "")
                .replace(/\s+/g, "");
        };

        for (const row of data) {
            const values = headers.map(h => {
                const key = normalizeKey(h);
                let val = row[key] !== undefined ? row[key] : (row[h] || '');

                val = String(val).replace(/"/g, '""');
                if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                    val = `"${val}"`;
                }
                return val;
            });
            rows.push(values.join(','));
        }
        return rows.join('\n');
    }

    let csvContent = '';

    if (preProdData.length > 0) {
        csvContent += '# COMPARAÇÃO PREPROD x PROD\n';
        csvContent += arrayToCSV(preProdData, ['Aplicação', 'Versão PREPROD', 'Versão PROD', 'Status', 'Path PREPROD', 'Path PROD']);
        csvContent += '\n\n';
    }

    if (preHmlData.length > 0) {
        csvContent += '# COMPARAÇÃO PREPROD x HML\n';
        csvContent += arrayToCSV(preHmlData, ['Aplicação', 'Versão PREPROD', 'Versão HML', 'Status', 'Path PREPROD', 'Path HML']);
        csvContent += '\n\n';
    }

    if (hmlProdData.length > 0) {
        csvContent += '# COMPARAÇÃO HML x PROD\n';
        csvContent += arrayToCSV(hmlProdData, ['Aplicação', 'Versão HML', 'Versão PROD', 'Status', 'Path HML', 'Path PROD']);
    }

    const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = `comparacao_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
});

// Gerenciador de Abas de Resultados
document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        tab.classList.add('active');
        document.getElementById(tab.dataset.tab).classList.add('active');
    });
});