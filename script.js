/* =========================================
   DECISÕES REFORMADAS - SCRIPT PRINCIPAL
   Gabinete Des. Alexandre Morais da Rosa
   ========================================= */

// =========================================
// VARIÁVEIS GLOBAIS
// =========================================

let dadosExcel = {}; // Dados de todas as planilhas
let planilhaAtual = null; // Nome da planilha atualmente selecionada
let dadosFiltrados = []; // Dados filtrados da planilha atual
let paginaAtual = 1;
let itensPorPagina = "all"; // Mostrar todos os registros sem paginação
let ultimaAtualizacao = null;
let filtrosPorColuna = {}; // Filtros individuais por coluna
let colunasVisiveis = {}; // Colunas visíveis por planilha
let dadosCelulas = []; // Armazena dados das células para clique

// Configurações padrão
let configuracoes = {
    tema: "dark",
    corDestaque: "blue",
    caminhoArquivo:
        "C:\\Apps\\DecisoesReformadas\\Reformadas TJSC - Resultado Agrupado.xlsx",
    itensPorPagina: 50,
    carregarAutomaticamente: true,
};

// =========================================
// INICIALIZAÇÃO
// =========================================

const ARQUIVO_XLSX_URL = '/resultadoagrupado.xlsx';

document.addEventListener("DOMContentLoaded", function () {
    carregarConfiguracoes();
    aplicarTema();

    // Event listeners para teclas de atalho
    document.addEventListener("keydown", handleKeyboardShortcuts);

    // Carregar dados automaticamente
    carregarArquivo();
});

// =========================================
// NAVEGAÇÃO
// =========================================

function showPage(pagina) {
    // Esconder todas as páginas
    document.querySelectorAll(".page-content").forEach((page) => {
        page.classList.add("hidden");
    });

    // Mostrar a página selecionada
    const pageElement = document.getElementById(`${pagina}-page`);
    if (pageElement) {
        pageElement.classList.remove("hidden");
    }

    // Atualizar navegação ativa
    document.querySelectorAll(".nav-item").forEach((item) => {
        item.classList.remove("active");
        if (item.dataset.page === pagina) {
            item.classList.add("active");
        }
    });
}

// =========================================
// CARREGAMENTO DE DADOS
// =========================================

async function carregarArquivo() {
    mostrarLoading("Carregando dados do servidor...");

    try {
        const response = await fetch(ARQUIVO_XLSX_URL);
        
        if (!response.ok) {
            throw new Error(`Erro ao carregar arquivo: ${response.status}`);
        }
        
        const dados = await response.arrayBuffer();
        await processarDadosExcel(dados);
        
    } catch (erro) {
        console.error("Erro ao carregar arquivo:", erro);
        mostrarToast("error", "Erro", "Não foi possível carregar os dados do servidor.");
        esconderLoading();
    }
}

async function processarDadosExcel(dados) {
    mostrarLoading("Processando planilhas...");

    try {
        const workbook = XLSX.read(dados, { type: "array", cellStyles: true });

        dadosExcel = {};

        // Processar cada planilha
        for (const nomeSheet of workbook.SheetNames) {
            const sheet = workbook.Sheets[nomeSheet];
            const dadosSheet = XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                defval: "",
            });

            if (dadosSheet.length > 0) {
                dadosExcel[nomeSheet] = {
                    cabecalhos: dadosSheet[0] || [],
                    dados: dadosSheet
                        .slice(1)
                        .filter((row) =>
                            row.some(
                                (cell) =>
                                    cell !== "" &&
                                    cell !== null &&
                                    cell !== undefined,
                            ),
                        ),
                };
            }
        }

        ultimaAtualizacao = new Date();

        // Atualizar interface
        atualizarKPIs();
        atualizarOverviewPlanilhas();
        atualizarEstatisticas();
        atualizarAcessoRapido();
        atualizarAbas();

        // Selecionar primeira planilha
        const primeiraSheet = Object.keys(dadosExcel)[0];
        if (primeiraSheet) {
            selecionarPlanilha(primeiraSheet);
        }

        mostrarToast(
            "success",
            "Sucesso!",
            `${Object.keys(dadosExcel).length} planilhas carregadas.`,
        );
    } catch (erro) {
        console.error("Erro ao processar arquivo:", erro);
        mostrarToast(
            "error",
            "Erro",
            "Não foi possível processar o arquivo Excel.",
        );
    } finally {
        esconderLoading();
    }
}

// =========================================
// ATUALIZAÇÃO DA INTERFACE
// =========================================

function atualizarKPIs() {
    // Total de registros
    let totalRegistros = 0;
    Object.values(dadosExcel).forEach((sheet) => {
        totalRegistros += sheet.dados.length;
    });

    document.getElementById("kpi-total").textContent =
        totalRegistros.toLocaleString("pt-BR");

    // Planilhas carregadas
    document.getElementById("kpi-planilhas").textContent =
        Object.keys(dadosExcel).length;

    // Última atualização
    if (ultimaAtualizacao) {
        document.getElementById("kpi-atualizacao").textContent =
            ultimaAtualizacao.toLocaleString("pt-BR", {
                day: "2-digit",
                month: "2-digit",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit",
            });
    }

    // Total de colunas
    let totalColunas = 0;
    Object.values(dadosExcel).forEach((sheet) => {
        totalColunas = Math.max(totalColunas, sheet.cabecalhos.length);
    });
    document.getElementById("kpi-colunas").textContent = totalColunas;
}

function atualizarOverviewPlanilhas() {
    const container = document.getElementById("sheets-overview");

    if (Object.keys(dadosExcel).length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <span class="material-symbols-outlined">upload_file</span>
                <p>Nenhum dado carregado</p>
                <small>Clique em "Atualizar Dados" para carregar o arquivo Excel</small>
            </div>
        `;
        return;
    }

    container.innerHTML = Object.entries(dadosExcel)
        .map(
            ([nome, sheet]) => `
        <div class="sheet-item" onclick="selecionarPlanilhaENavegar('${nome}')">
            <div class="sheet-item-info">
                <div class="sheet-item-icon">
                    <span class="material-symbols-outlined">table_chart</span>
                </div>
                <div>
                    <div class="sheet-item-name">${nome}</div>
                    <div class="sheet-item-count">${sheet.dados.length.toLocaleString("pt-BR")} registros • ${sheet.cabecalhos.length} colunas</div>
                </div>
            </div>
            <span class="material-symbols-outlined" style="color: var(--text-tertiary);">chevron_right</span>
        </div>
    `,
        )
        .join("");
}

function atualizarEstatisticas() {
    const container = document.getElementById("stats-chart");

    if (Object.keys(dadosExcel).length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <span class="material-symbols-outlined">insights</span>
                <p>Estatísticas indisponíveis</p>
                <small>Carregue os dados para visualizar</small>
            </div>
        `;
        return;
    }

    // Encontrar o máximo para calcular porcentagens
    let maxRegistros = 0;
    Object.values(dadosExcel).forEach((sheet) => {
        maxRegistros = Math.max(maxRegistros, sheet.dados.length);
    });

    container.innerHTML = Object.entries(dadosExcel)
        .map(([nome, sheet]) => {
            const porcentagem = (sheet.dados.length / maxRegistros) * 100;
            return `
            <div class="stat-bar">
                <span class="stat-bar-label" title="${nome}">${nome}</span>
                <div class="stat-bar-track">
                    <div class="stat-bar-fill" style="width: ${porcentagem}%">
                        <span class="stat-bar-value">${sheet.dados.length.toLocaleString("pt-BR")}</span>
                    </div>
                </div>
            </div>
        `;
        })
        .join("");
}

function atualizarAcessoRapido() {
    const container = document.getElementById("quick-access");

    if (Object.keys(dadosExcel).length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <span class="material-symbols-outlined">touch_app</span>
                <p>Acesso rápido</p>
                <small>As planilhas carregadas aparecerão aqui para acesso direto</small>
            </div>
        `;
        return;
    }

    container.innerHTML = Object.keys(dadosExcel)
        .map(
            (nome) => `
        <div class="quick-access-item" onclick="selecionarPlanilhaENavegar('${nome}')">
            <span class="material-symbols-outlined">table_view</span>
            <span>${nome}</span>
        </div>
    `,
        )
        .join("");
}

function atualizarAbas() {
    const container = document.getElementById("sheet-tabs");

    if (Object.keys(dadosExcel).length === 0) {
        container.innerHTML =
            '<div class="empty-tab">Nenhuma planilha carregada</div>';
        return;
    }

    container.innerHTML = Object.keys(dadosExcel)
        .map(
            (nome) => `
        <div class="sheet-tab ${nome === planilhaAtual ? "active" : ""}" 
             onclick="selecionarPlanilha('${nome}')"
             data-sheet="${nome}">
            <span class="material-symbols-outlined">table_chart</span>
            ${nome}
        </div>
    `,
        )
        .join("");
}

// =========================================
// SELEÇÃO E RENDERIZAÇÃO DE PLANILHAS
// =========================================

function selecionarPlanilha(nome) {
    if (!dadosExcel[nome]) return;

    planilhaAtual = nome;
    paginaAtual = 1;

    // Atualizar abas
    document.querySelectorAll(".sheet-tab").forEach((tab) => {
        tab.classList.remove("active");
        if (tab.dataset.sheet === nome) {
            tab.classList.add("active");
        }
    });

    // Limpar filtros
    document.getElementById("table-filter").value = "";
    filtrosPorColuna = {};

    // Carregar dados
    dadosFiltrados = [...dadosExcel[nome].dados];

    // Renderizar tabela
    renderizarTabela();
}

function selecionarPlanilhaENavegar(nome) {
    selecionarPlanilha(nome);
    showPage("tabelas");
}

function renderizarTabela() {
    const container = document.getElementById("table-container");

    if (!planilhaAtual || !dadosExcel[planilhaAtual]) {
        container.innerHTML = `
            <div class="empty-state large">
                <span class="material-symbols-outlined">table_rows</span>
                <p>Nenhuma tabela selecionada</p>
                <small>Selecione uma planilha nas abas acima</small>
            </div>
        `;
        return;
    }

    const sheet = dadosExcel[planilhaAtual];
    const cabecalhos = sheet.cabecalhos;

    // Inicializar colunas visíveis (ocultar link_decisao por padrão)
    if (!colunasVisiveis[planilhaAtual]) {
        colunasVisiveis[planilhaAtual] = cabecalhos
            .map((col, i) => ({ col: col?.toLowerCase(), index: i }))
            .filter((item) => item.col !== "link_decisao")
            .map((item) => item.index);
    }
    const cabecalhosVisiveis = colunasVisiveis[planilhaAtual];

    // Mostrar todos os registros
    const totalRegistros = dadosFiltrados.length;
    const inicio = 0;
    const fim = totalRegistros;
    const dadosPagina = dadosFiltrados.slice(inicio, fim);

    // Gerar tabela
    let html = `
        <table class="data-table">
            <thead>
                <tr>
                    <th class="row-number">#</th>
                    ${cabecalhosVisiveis.map((colIndex) => `<th>${escapeHtml(cabecalhos[colIndex] || "")}</th>`).join("")}
                </tr>
                <tr class="filter-row">
                    <th class="row-number">
                        <button class="btn-clear-filters" onclick="limparFiltrosColunas()" title="Limpar todos os filtros">
                            <span class="material-symbols-outlined">filter_alt_off</span>
                        </button>
                    </th>
                    ${cabecalhosVisiveis
                        .map(
                            (colIndex) => `
                        <th class="filter-cell">
                            <input type="text" 
                                   class="column-filter-input" 
                                   placeholder="Filtrar..." 
                                   data-column="${colIndex}"
                                   value="${escapeHtml(filtrosPorColuna[colIndex] || "")}"
                                   onkeyup="filtrarPorColuna(event, ${colIndex})"
                                   oninput="filtrarPorColunaDebounce(${colIndex}, this.value)">
                        </th>
                    `,
                        )
                        .join("")}
                </tr>
            </thead>
            <tbody>
    `;

    if (dadosPagina.length === 0) {
        html += `
            <tr>
                <td colspan="${cabecalhosVisiveis.length + 1}" style="text-align: center; padding: 40px; color: var(--text-tertiary);">
                    Nenhum registro encontrado
                </td>
            </tr>
        `;
    } else {
        dadosCelulas = []; // Limpar dados das células
        dadosPagina.forEach((row, index) => {
            const numeroLinha = inicio + index + 1;
            html += `<tr>
                <td class="row-number">${numeroLinha}</td>
                ${cabecalhosVisiveis
                    .map((colIndex) => {
                        const valor =
                            row[colIndex] !== undefined ? row[colIndex] : "";
                        let valorStr = String(valor);
                        const nomeColuna = (
                            cabecalhos[colIndex] || ""
                        ).toLowerCase();

                        // Formatação especial para coluna data_decisao
                        let valorExibicao = valorStr;
                        if (nomeColuna === "data_decisao") {
                            valorExibicao = extrairData(valorStr);
                        }

                        // Coluna link_pdf_download: abre link diretamente
                        if (
                            nomeColuna === "link_pdf_download" &&
                            valorStr.trim()
                        ) {
                            return `<td class="link-cell" 
                                   onclick="abrirLink('${escapeHtml(valorStr.trim()).replace(/'/g, "\\'")}')" 
                                   title="Clique para abrir o PDF">
                                   <span class="material-symbols-outlined">open_in_new</span>
                                   Abrir PDF
                               </td>`;
                        }

                        const valorTruncado =
                            valorExibicao.length > 100
                                ? valorExibicao.substring(0, 100) + "..."
                                : valorExibicao;
                        const clickable =
                            valorStr.length > 50 ? "clickable" : "";
                        // Armazenar dados da célula para acesso via índice
                        const celulaIndex = dadosCelulas.length;
                        dadosCelulas.push({
                            valor: valorStr,
                            coluna:
                                cabecalhos[colIndex] ||
                                "Coluna " + (colIndex + 1),
                        });
                        return `<td class="${clickable}" 
                               ${clickable ? `data-celula-index="${celulaIndex}" onclick="mostrarDetalhesCelulaIndex(${celulaIndex})"` : ""}
                               title="${escapeHtml(valorStr)}">${escapeHtml(valorTruncado)}</td>`;
                    })
                    .join("")}
            </tr>`;
        });
    }

    html += "</tbody></table>";
    container.innerHTML = html;

    // Atualizar informações
    document.getElementById("filter-count").textContent =
        `${totalRegistros.toLocaleString("pt-BR")} registros`;

    // Esconder paginação (mostrando todos)
    document.getElementById("pagination-container").style.display = "none";
}

function renderizarPaginacao(totalPaginas) {
    const container = document.getElementById("pagination-pages");
    const btnPrev = document.getElementById("btn-prev");
    const btnNext = document.getElementById("btn-next");

    // Atualizar botões prev/next
    btnPrev.disabled = paginaAtual === 1;
    btnNext.disabled = paginaAtual === totalPaginas || totalPaginas === 0;

    // Gerar números de página
    let html = "";
    const maxPaginas = 5;
    let inicioPag = Math.max(1, paginaAtual - Math.floor(maxPaginas / 2));
    let fimPag = Math.min(totalPaginas, inicioPag + maxPaginas - 1);

    if (fimPag - inicioPag < maxPaginas - 1) {
        inicioPag = Math.max(1, fimPag - maxPaginas + 1);
    }

    if (inicioPag > 1) {
        html += `<span class="page-number" onclick="irParaPagina(1)">1</span>`;
        if (inicioPag > 2) {
            html += `<span class="page-number" style="cursor: default;">...</span>`;
        }
    }

    for (let i = inicioPag; i <= fimPag; i++) {
        html += `<span class="page-number ${i === paginaAtual ? "active" : ""}" onclick="irParaPagina(${i})">${i}</span>`;
    }

    if (fimPag < totalPaginas) {
        if (fimPag < totalPaginas - 1) {
            html += `<span class="page-number" style="cursor: default;">...</span>`;
        }
        html += `<span class="page-number" onclick="irParaPagina(${totalPaginas})">${totalPaginas}</span>`;
    }

    container.innerHTML = html;
}

// =========================================
// PAGINAÇÃO
// =========================================

function irParaPagina(pagina) {
    paginaAtual = pagina;
    renderizarTabela();

    // Scroll para o topo da tabela
    document.getElementById("table-container").scrollTop = 0;
}

function paginaAnterior() {
    if (paginaAtual > 1) {
        irParaPagina(paginaAtual - 1);
    }
}

function proximaPagina() {
    const totalPaginas =
        itensPorPagina === "all"
            ? 1
            : Math.ceil(dadosFiltrados.length / itensPorPagina);
    if (paginaAtual < totalPaginas) {
        irParaPagina(paginaAtual + 1);
    }
}

function alterarTamanhoPagina() {
    const select = document.getElementById("page-size");
    itensPorPagina = select.value === "all" ? "all" : parseInt(select.value);
    paginaAtual = 1;
    renderizarTabela();
}

// =========================================
// FILTRO DE TABELA
// =========================================

function filtrarTabelaAtual() {
    if (!planilhaAtual || !dadosExcel[planilhaAtual]) return;

    // Usa a função unificada de filtros
    aplicarFiltrosColunas();
}

function limparFiltroTabela() {
    document.getElementById("table-filter").value = "";
    filtrarTabelaAtual();
}

// =========================================
// FILTRO POR COLUNA
// =========================================

let filtroDebounceTimer = null;

function filtrarPorColuna(event, colIndex) {
    // Se pressionar Enter, aplica o filtro imediatamente
    if (event.key === "Enter") {
        clearTimeout(filtroDebounceTimer);
        aplicarFiltrosColunas();
    }
}

function filtrarPorColunaDebounce(colIndex, valor) {
    // Atualizar valor do filtro
    if (valor.trim()) {
        filtrosPorColuna[colIndex] = valor;
    } else {
        delete filtrosPorColuna[colIndex];
    }

    // Debounce para não filtrar a cada tecla
    clearTimeout(filtroDebounceTimer);
    filtroDebounceTimer = setTimeout(() => {
        aplicarFiltrosColunas();
    }, 300);
}

function aplicarFiltrosColunas() {
    if (!planilhaAtual || !dadosExcel[planilhaAtual]) return;

    const sheet = dadosExcel[planilhaAtual];
    const filtroGeral = document
        .getElementById("table-filter")
        .value.toLowerCase();

    // Começar com todos os dados
    dadosFiltrados = sheet.dados.filter((row) => {
        // Primeiro aplicar filtro geral (se houver)
        if (filtroGeral) {
            const matchGeral = row.some((cell) =>
                String(cell).toLowerCase().includes(filtroGeral),
            );
            if (!matchGeral) return false;
        }

        // Depois aplicar filtros por coluna
        for (const [colIndex, filtro] of Object.entries(filtrosPorColuna)) {
            const valor = String(row[parseInt(colIndex)] || "").toLowerCase();
            const filtroLower = filtro.toLowerCase();

            if (!valor.includes(filtroLower)) {
                return false;
            }
        }

        return true;
    });

    paginaAtual = 1;
    renderizarTabela();

    // Restaurar foco no input ativo (se houver)
    const activeElement = document.activeElement;
    if (
        activeElement &&
        activeElement.classList.contains("column-filter-input")
    ) {
        const colIndex = activeElement.dataset.column;
        setTimeout(() => {
            const input = document.querySelector(
                `.column-filter-input[data-column="${colIndex}"]`,
            );
            if (input) {
                input.focus();
                // Colocar cursor no final do texto
                const len = input.value.length;
                input.setSelectionRange(len, len);
            }
        }, 0);
    }
}

function limparFiltrosColunas() {
    filtrosPorColuna = {};
    aplicarFiltrosColunas();
    mostrarToast(
        "info",
        "Filtros limpos",
        "Todos os filtros de coluna foram removidos.",
    );
}

// =========================================
// SELETOR DE COLUNAS
// =========================================

let colunasTemporarias = [];

function abrirSeletorColunas() {
    if (!planilhaAtual || !dadosExcel[planilhaAtual]) {
        mostrarToast("warning", "Atenção", "Selecione uma planilha primeiro.");
        return;
    }

    const sheet = dadosExcel[planilhaAtual];
    const cabecalhos = sheet.cabecalhos;

    // Inicializar colunas visíveis se não existir
    if (!colunasVisiveis[planilhaAtual]) {
        colunasVisiveis[planilhaAtual] = cabecalhos.map((_, i) => i);
    }

    // Cópia temporária para edição
    colunasTemporarias = [...colunasVisiveis[planilhaAtual]];

    // Renderizar lista de colunas
    const container = document.getElementById("columns-list");
    container.innerHTML = cabecalhos
        .map((col, index) => {
            const checked = colunasTemporarias.includes(index) ? "checked" : "";
            const nomeColuna = col || `Coluna ${index + 1}`;
            return `
            <label class="column-checkbox">
                <input type="checkbox" 
                       value="${index}" 
                       ${checked}
                       onchange="toggleColunaTemporaria(${index}, this.checked)">
                <span class="checkmark"></span>
                <span class="column-name">${escapeHtml(nomeColuna)}</span>
            </label>
        `;
        })
        .join("");

    document.getElementById("columns-modal").classList.add("active");
}

function fecharModalColunas() {
    document.getElementById("columns-modal").classList.remove("active");
}

function toggleColunaTemporaria(index, checked) {
    if (checked) {
        if (!colunasTemporarias.includes(index)) {
            colunasTemporarias.push(index);
            colunasTemporarias.sort((a, b) => a - b);
        }
    } else {
        colunasTemporarias = colunasTemporarias.filter((i) => i !== index);
    }
}

function selecionarTodasColunas() {
    if (!planilhaAtual || !dadosExcel[planilhaAtual]) return;

    const sheet = dadosExcel[planilhaAtual];
    colunasTemporarias = sheet.cabecalhos.map((_, i) => i);

    // Atualizar checkboxes
    document
        .querySelectorAll("#columns-list input[type='checkbox']")
        .forEach((cb) => {
            cb.checked = true;
        });
}

function deselecionarTodasColunas() {
    colunasTemporarias = [];

    // Atualizar checkboxes
    document
        .querySelectorAll("#columns-list input[type='checkbox']")
        .forEach((cb) => {
            cb.checked = false;
        });
}

function aplicarSelecaoColunas() {
    if (colunasTemporarias.length === 0) {
        mostrarToast("warning", "Atenção", "Selecione pelo menos uma coluna.");
        return;
    }

    colunasVisiveis[planilhaAtual] = [...colunasTemporarias];
    fecharModalColunas();
    renderizarTabela();

    const total = dadosExcel[planilhaAtual].cabecalhos.length;
    const visiveis = colunasVisiveis[planilhaAtual].length;
    mostrarToast(
        "success",
        "Colunas atualizadas",
        `${visiveis} de ${total} colunas visíveis.`,
    );
}

// =========================================
// PESQUISA GLOBAL
// =========================================

function handleSearchKeyup(event) {
    if (event.key === "Enter") {
        realizarPesquisa();
    }
}

function realizarPesquisa() {
    const termo = document.getElementById("global-search").value.trim();

    if (!termo) {
        mostrarToast("warning", "Atenção", "Digite um termo para pesquisar.");
        return;
    }

    if (Object.keys(dadosExcel).length === 0) {
        mostrarToast(
            "warning",
            "Atenção",
            "Carregue os dados antes de pesquisar.",
        );
        return;
    }

    const caseSensitive = document.getElementById(
        "search-case-sensitive",
    ).checked;
    const exactMatch = document.getElementById("search-exact-match").checked;
    const allSheets = document.getElementById("search-all-sheets").checked;

    const resultados = [];
    const termoProcessado = caseSensitive ? termo : termo.toLowerCase();

    const sheetsParaBuscar = allSheets
        ? Object.keys(dadosExcel)
        : planilhaAtual
          ? [planilhaAtual]
          : [];

    sheetsParaBuscar.forEach((nomeSheet) => {
        const sheet = dadosExcel[nomeSheet];

        sheet.dados.forEach((row, rowIndex) => {
            row.forEach((cell, colIndex) => {
                const cellStr = String(cell);
                const cellProcessada = caseSensitive
                    ? cellStr
                    : cellStr.toLowerCase();

                let encontrado = false;
                if (exactMatch) {
                    encontrado = cellProcessada === termoProcessado;
                } else {
                    encontrado = cellProcessada.includes(termoProcessado);
                }

                if (encontrado) {
                    resultados.push({
                        sheet: nomeSheet,
                        linha: rowIndex + 1,
                        coluna:
                            sheet.cabecalhos[colIndex] ||
                            `Coluna ${colIndex + 1}`,
                        colunaIndex: colIndex,
                        valor: cellStr,
                        termo: termo,
                    });
                }
            });
        });
    });

    renderizarResultadosPesquisa(resultados, termo);
}

function renderizarResultadosPesquisa(resultados, termo) {
    const container = document.getElementById("search-results");
    const countElement = document.getElementById("results-count");

    countElement.textContent = `${resultados.length.toLocaleString("pt-BR")} resultado${resultados.length !== 1 ? "s" : ""} encontrado${resultados.length !== 1 ? "s" : ""}`;

    if (resultados.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <span class="material-symbols-outlined">search_off</span>
                <p>Nenhum resultado encontrado</p>
                <small>Tente ajustar os termos ou opções de pesquisa</small>
            </div>
        `;
        return;
    }

    // Limitar a 100 resultados para performance
    const resultadosLimitados = resultados.slice(0, 100);

    container.innerHTML = resultadosLimitados
        .map((resultado) => {
            const valorDestacado = destacarTermo(
                resultado.valor,
                resultado.termo,
            );
            return `
            <div class="search-result-item" onclick="navegarParaResultado('${resultado.sheet}', ${resultado.linha})">
                <div class="result-header">
                    <span class="result-sheet">
                        <span class="material-symbols-outlined" style="font-size: 14px;">table_chart</span>
                        ${resultado.sheet}
                    </span>
                    <span class="result-location">Linha ${resultado.linha} • ${resultado.coluna}</span>
                </div>
                <div class="result-content">${valorDestacado}</div>
            </div>
        `;
        })
        .join("");

    if (resultados.length > 100) {
        container.innerHTML += `
            <div style="padding: 20px; text-align: center; color: var(--text-tertiary);">
                Mostrando 100 de ${resultados.length.toLocaleString("pt-BR")} resultados
            </div>
        `;
    }
}

function destacarTermo(texto, termo) {
    const textoStr = String(texto);
    const regex = new RegExp(`(${escapeRegex(termo)})`, "gi");
    return escapeHtml(textoStr).replace(regex, "<mark>$1</mark>");
}

function navegarParaResultado(nomeSheet, linha) {
    selecionarPlanilha(nomeSheet);
    showPage("tabelas");

    // Calcular a página onde está o resultado
    const paginaDestino = Math.ceil(
        linha /
            (itensPorPagina === "all" ? dadosFiltrados.length : itensPorPagina),
    );
    irParaPagina(paginaDestino);

    mostrarToast(
        "info",
        "Navegação",
        `Ir para linha ${linha} na planilha "${nomeSheet}"`,
    );
}

// =========================================
// MODAL DE DETALHES
// =========================================

let conteudoModalAtual = "";

function mostrarDetalhesCelulaIndex(index) {
    if (dadosCelulas[index]) {
        const { valor, coluna } = dadosCelulas[index];
        mostrarDetalhesCelula(valor, coluna);
    }
}

function abrirLink(url) {
    if (!url) return;

    // Adiciona protocolo se não existir
    let urlFinal = url.trim();
    if (!urlFinal.startsWith("http://") && !urlFinal.startsWith("https://")) {
        urlFinal = "https://" + urlFinal;
    }

    window.open(urlFinal, "_blank", "noopener,noreferrer");
}

function mostrarDetalhesCelula(valor, coluna) {
    conteudoModalAtual = valor;
    document.getElementById("modal-title").textContent = coluna;
    document.getElementById("modal-body").innerHTML =
        `<pre>${escapeHtml(valor)}</pre>`;
    document.getElementById("cell-modal").classList.add("active");
}

function fecharModal() {
    document.getElementById("cell-modal").classList.remove("active");
}

function copiarConteudoModal() {
    navigator.clipboard
        .writeText(conteudoModalAtual)
        .then(() => {
            mostrarToast(
                "success",
                "Copiado!",
                "Conteúdo copiado para a área de transferência.",
            );
        })
        .catch(() => {
            mostrarToast(
                "error",
                "Erro",
                "Não foi possível copiar o conteúdo.",
            );
        });
}

// =========================================
// EXPORTAÇÃO
// =========================================

function exportarPlanilhaAtual(formato) {
    if (!planilhaAtual || !dadosExcel[planilhaAtual]) {
        mostrarToast(
            "warning",
            "Atenção",
            "Selecione uma planilha para exportar.",
        );
        return;
    }

    const sheet = dadosExcel[planilhaAtual];
    const dados = [sheet.cabecalhos, ...dadosFiltrados];

    if (formato === "excel") {
        exportarExcel(dados, planilhaAtual);
    } else if (formato === "pdf") {
        exportarPDF(dados, planilhaAtual);
    }
}

function exportarExcel(dados, nomeArquivo) {
    const ws = XLSX.utils.aoa_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, nomeArquivo);

    XLSX.writeFile(wb, `${nomeArquivo}_${formatarDataArquivo()}.xlsx`);
    mostrarToast("success", "Exportado!", "Arquivo Excel gerado com sucesso.");
}

function exportarPDF(dados, nomeArquivo) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF("landscape");

    // Título
    doc.setFillColor(59, 130, 246);
    doc.rect(0, 0, doc.internal.pageSize.width, 25, "F");

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont("helvetica", "bold");
    doc.text(`Decisões Reformadas - ${nomeArquivo}`, 15, 16);

    doc.setTextColor(59, 130, 246);
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.text(`Gerado em: ${new Date().toLocaleString("pt-BR")}`, 15, 35);
    doc.text(`Total de registros: ${dados.length - 1}`, 15, 42);

    // Tabela
    const cabecalhos = dados[0];
    const corpo = dados.slice(1);

    doc.autoTable({
        head: [cabecalhos],
        body: corpo,
        startY: 50,
        styles: {
            fontSize: 7,
            cellPadding: 2,
            overflow: "linebreak",
            cellWidth: "wrap",
        },
        headStyles: {
            fillColor: [59, 130, 246],
            textColor: [255, 255, 255],
            fontStyle: "bold",
        },
        alternateRowStyles: {
            fillColor: [248, 250, 252],
        },
    });

    doc.save(`${nomeArquivo}_${formatarDataArquivo()}.pdf`);
    mostrarToast("success", "Exportado!", "Arquivo PDF gerado com sucesso.");
}

function formatarDataArquivo() {
    return new Date().toISOString().split("T")[0];
}

// =========================================
// TEMA
// =========================================

function toggleTheme() {
    const temas = ["dark", "light"];
    const temaAtual =
        document.documentElement.getAttribute("data-theme") || "dark";
    const proximoIndex = (temas.indexOf(temaAtual) + 1) % temas.length;
    const proximoTema = temas[proximoIndex];

    document.documentElement.setAttribute("data-theme", proximoTema);
    configuracoes.tema = proximoTema;
    salvarConfiguracoes();

    mostrarToast(
        "info",
        "Tema alterado",
        `Tema ${proximoTema === "dark" ? "escuro" : "claro"} ativado.`,
    );
}

function aplicarTema() {
    document.documentElement.setAttribute(
        "data-theme",
        configuracoes.tema || "dark",
    );
    document.documentElement.setAttribute(
        "data-accent",
        configuracoes.corDestaque || "blue",
    );

    // Atualizar selects de configuração se existirem
    const selectTema = document.getElementById("config-theme");
    const selectAccent = document.getElementById("config-accent");

    if (selectTema) selectTema.value = configuracoes.tema || "dark";
    if (selectAccent) selectAccent.value = configuracoes.corDestaque || "blue";
}

// =========================================
// CONFIGURAÇÕES
// =========================================

function carregarConfiguracoes() {
    const configSalvas = localStorage.getItem("decisoesReformadas_config");
    if (configSalvas) {
        configuracoes = { ...configuracoes, ...JSON.parse(configSalvas) };
    }

    // Aplicar configurações
    itensPorPagina = configuracoes.itensPorPagina || 50;

    // Atualizar campos de configuração
    setTimeout(() => {
        const autoloadCheckbox = document.getElementById("config-autoload");
        const pageSizeSelect = document.getElementById("config-pagesize");

        if (autoloadCheckbox)
            autoloadCheckbox.checked = configuracoes.carregarAutomaticamente;
        if (pageSizeSelect) pageSizeSelect.value = configuracoes.itensPorPagina;
    }, 100);
}

function salvarConfiguracoes() {
    // Coletar valores dos campos
    const selectTema = document.getElementById("config-theme");
    const selectAccent = document.getElementById("config-accent");
    const selectPageSize = document.getElementById("config-pagesize");
    const checkboxAutoload = document.getElementById("config-autoload");

    if (selectTema) configuracoes.tema = selectTema.value;
    if (selectAccent) configuracoes.corDestaque = selectAccent.value;
    if (selectPageSize)
        configuracoes.itensPorPagina = parseInt(selectPageSize.value);
    if (checkboxAutoload)
        configuracoes.carregarAutomaticamente = checkboxAutoload.checked;

    // Aplicar tema
    aplicarTema();

    // Salvar no localStorage
    localStorage.setItem(
        "decisoesReformadas_config",
        JSON.stringify(configuracoes),
    );
}

// =========================================
// UTILITÁRIOS
// =========================================

function escapeHtml(texto) {
    const div = document.createElement("div");
    div.textContent = texto;
    return div.innerHTML;
}

function extrairData(texto) {
    // Extrai apenas a data de strings como:
    // "2023/0046924-0 de 25/02/2025" -> "25/02/2025"
    // "2025/0020542-7 - 24/04/2025" -> "24/04/2025"
    // Se já for uma data pura, retorna ela mesma
    if (!texto) return texto;

    // Regex para encontrar data no formato DD/MM/YYYY
    const regexData = /(\d{2}\/\d{2}\/\d{4})/;
    const match = texto.match(regexData);

    if (match) {
        return match[1];
    }

    // Tenta formato YYYY-MM-DD
    const regexDataISO = /(\d{4}-\d{2}-\d{2})/;
    const matchISO = texto.match(regexDataISO);

    if (matchISO) {
        // Converte de YYYY-MM-DD para DD/MM/YYYY
        const [ano, mes, dia] = matchISO[1].split("-");
        return `${dia}/${mes}/${ano}`;
    }

    return texto;
}

function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function mostrarLoading(mensagem = "Carregando...") {
    const overlay = document.getElementById("loading-overlay");
    const texto = overlay.querySelector(".loading-text");
    texto.textContent = mensagem;
    overlay.classList.add("active");
}

function esconderLoading() {
    document.getElementById("loading-overlay").classList.remove("active");
}

function mostrarToast(tipo, titulo, mensagem) {
    const container = document.getElementById("toast-container");

    const icones = {
        success: "check_circle",
        error: "error",
        warning: "warning",
        info: "info",
    };

    const toast = document.createElement("div");
    toast.className = `toast ${tipo}`;
    toast.innerHTML = `
        <div class="toast-icon">
            <span class="material-symbols-outlined">${icones[tipo]}</span>
        </div>
        <div class="toast-content">
            <div class="toast-title">${titulo}</div>
            <div class="toast-message">${mensagem}</div>
        </div>
        <button class="toast-close" onclick="this.parentElement.remove()">
            <span class="material-symbols-outlined">close</span>
        </button>
    `;

    container.appendChild(toast);

    // Auto-remover após 5 segundos
    setTimeout(() => {
        toast.classList.add("toast-out");
        setTimeout(() => toast.remove(), 300);
    }, 5000);
}

function handleKeyboardShortcuts(event) {
    // Ctrl + O - Abrir arquivo
    if (event.ctrlKey && event.key === "o") {
        event.preventDefault();
        carregarArquivo();
    }

    // Ctrl + F - Ir para pesquisa
    if (event.ctrlKey && event.key === "f") {
        event.preventDefault();
        showPage("pesquisa");
        document.getElementById("global-search").focus();
    }

    // Escape - Fechar modal
    if (event.key === "Escape") {
        fecharModal();
    }
}

// =========================================
// FECHAMENTO DE MODAL AO CLICAR FORA
// =========================================

document.addEventListener("click", function (event) {
    const modal = document.getElementById("cell-modal");
    if (event.target === modal) {
        fecharModal();
    }
});
