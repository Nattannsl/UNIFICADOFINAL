const STORAGE_KEY = "gw_khomp_bases_v1";

let state = {
    bases: {},
    baseAtiva: "",
    filtroEquip: "TODOS",
    busca: "",
    arquivoPendente: null
};

const baseSelect = document.getElementById("baseSelect");
const nomeBaseInput = document.getElementById("nomeBaseInput");
const fileInput = document.getElementById("fileInput");
const fileDrop = document.getElementById("fileDrop");
const fileNameLabel = document.getElementById("fileNameLabel");
const btnSubirPlanilha = document.getElementById("btnSubirPlanilha");
const btnExcluirBase = document.getElementById("btnExcluirBase");
const btnLimparFiltros = document.getElementById("btnLimparFiltros");
const searchInput = document.getElementById("searchInput");
const totalRegistrosEl = document.getElementById("totalRegistros");
const cardsContainer = document.getElementById("cardsContainer");

function salvarStorage() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
        bases: state.bases,
        baseAtiva: state.baseAtiva
    }));
}

function carregarStorage() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
        try {
            const parsed = JSON.parse(raw);
            state.bases = parsed.bases || {};
            state.baseAtiva = parsed.baseAtiva || "";
        } catch (e) {
            console.error("Erro ao ler storage:", e);
        }
    }

    const nomes = Object.keys(state.bases);
    if (!state.baseAtiva && nomes.length > 0) {
        state.baseAtiva = nomes[0];
    }
}

function normalizarChave(str) {
    return String(str || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, "");
}

function valorValido(val) {
    if (val === undefined || val === null) return false;
    const s = String(val).trim();
    return s !== "" && s.toUpperCase() !== "N/U" && s !== "-";
}

function extrairRegistrosDaPlanilha(rows) {
    const registros = [];

    for (const row of rows) {
        const mapa = {};
        for (const [k, v] of Object.entries(row)) {
            mapa[normalizarChave(k)] = v;
        }

        const empresa = String(
            mapa["EMPRESA"] ?? mapa["CLIENTE"] ?? mapa["NOMEDOCLIENTE"] ?? ""
        ).trim();

        const codigo = String(
            mapa["COD"] ?? mapa["CODIGO"] ?? ""
        ).trim();

        const gwIp = String(
            mapa["GWPORTA"] ?? mapa["GW"] ?? mapa["IPGW"] ?? ""
        ).trim();

        const gwSenha = String(
            mapa["SENHAGW"] ?? mapa["SENHA"] ?? ""
        ).trim();

        const khompIp = String(
            mapa["KHOMPPORTA"] ?? mapa["KHOMP"] ?? mapa["IPKHOMP"] ?? ""
        ).trim();

        const khompSenha = String(
            mapa["SENHAKHOMP"] ?? mapa["SENHAKH"] ?? mapa["SENHA"] ?? ""
        ).trim();

        if (!empresa && !codigo) continue;

        if (valorValido(gwIp)) {
            registros.push({
                codigo: codigo || "-",
                equipamento: "GW",
                empresa: empresa || "Sem Nome",
                ip: gwIp,
                senha: valorValido(gwSenha) ? gwSenha : "-"
            });
        }

        if (valorValido(khompIp)) {
            registros.push({
                codigo: codigo || "-",
                equipamento: "KHOMP",
                empresa: empresa || "Sem Nome",
                ip: khompIp,
                senha: valorValido(khompSenha) ? khompSenha : "-"
            });
        }
    }

    return registros;
}

function processarArquivo(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" });
                const primeiraAba = workbook.SheetNames[0];
                const sheet = workbook.Sheets[primeiraAba];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                resolve(extrairRegistrosDaPlanilha(rows));
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function renderSelectBases() {
    const nomes = Object.keys(state.bases);
    baseSelect.innerHTML = "";

    if (nomes.length === 0) {
        const opt = document.createElement("option");
        opt.value = "";
        opt.textContent = "Nenhuma base importada";
        baseSelect.appendChild(opt);
        return;
    }

    nomes.forEach((nome) => {
        const opt = document.createElement("option");
        opt.value = nome;
        opt.textContent = nome;
        if (nome === state.baseAtiva) opt.selected = true;
        baseSelect.appendChild(opt);
    });
}

function formatarBlocoCopia(item) {
    return [
        `CÓDIGO: ${item.codigo}`,
        `EMPRESA: ${item.empresa}`,
        `EQUIPAMENTO: ${item.equipamento}`,
        `IP / PORTA: ${item.ip}`,
        `SENHA: ${item.senha}`
    ].join("\n");
}

function renderCards() {
    const listaBase = state.bases[state.baseAtiva] || [];
    const termo = state.busca.trim().toLowerCase();

    const filtrados = listaBase.filter((item) => {
        const matchEquip =
            state.filtroEquip === "TODOS" || item.equipamento === state.filtroEquip;

        const matchBusca =
            !termo ||
            String(item.codigo).toLowerCase().includes(termo) ||
            String(item.empresa).toLowerCase().includes(termo) ||
            String(item.equipamento).toLowerCase().includes(termo) ||
            String(item.ip).toLowerCase().includes(termo) ||
            String(item.senha).toLowerCase().includes(termo);

        return matchEquip && matchBusca;
    });

    totalRegistrosEl.textContent = filtrados.length;
    cardsContainer.innerHTML = "";

    if (filtrados.length === 0) {
        cardsContainer.innerHTML = `
      <div class="empty-state">
        Nenhum registro encontrado. Importe uma planilha (.xlsx ou .csv) na barra lateral ou ajuste os filtros.
      </div>
    `;
        return;
    }

    filtrados.forEach((item) => {
        const card = document.createElement("div");
        card.className = "card";

        function desenharModoLeitura() {
            card.classList.remove("editing");
            const badgeClass = item.equipamento === "GW" ? "badge-gw" : "badge-khomp";

            card.innerHTML = `
        <div class="card-top">
          <button type="button" class="btn-lock-edit" title="Destravar para editar">
            <span>🔒</span>
            <span>Editar</span>
          </button>
        </div>

        <div class="card-body">
          <div>
            <div class="col-label">Código</div>
            <span class="badge-cod">${escapeHtml(item.codigo)}</span>
          </div>

          <div>
            <div class="col-label">Equip.</div>
            <span class="badge-equip ${badgeClass}">${escapeHtml(item.equipamento)}</span>
          </div>

          <div>
            <div class="col-label">Nome da Empresa</div>
            <div class="field-box field-empresa" title="${escapeAttr(item.empresa)}">${escapeHtml(item.empresa)}</div>
          </div>

          <div>
            <div class="col-label">IP / Porta</div>
            <div class="field-box" title="${escapeAttr(item.ip)}">${escapeHtml(item.ip)}</div>
          </div>

          <div>
            <div class="col-label">Senha</div>
            <div class="field-box" title="${escapeAttr(item.senha)}">${escapeHtml(item.senha)}</div>
          </div>

          <div>
            <button type="button" class="btn-copy">Copiar Bloco</button>
          </div>
        </div>
      `;

            card.querySelector(".btn-lock-edit").addEventListener("click", desenharModoEdicao);

            const btnCopy = card.querySelector(".btn-copy");
            btnCopy.addEventListener("click", async () => {
                const texto = formatarBlocoCopia(item);
                try {
                    await navigator.clipboard.writeText(texto);
                } catch {
                    const ta = document.createElement("textarea");
                    ta.value = texto;
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand("copy");
                    document.body.removeChild(ta);
                }
                btnCopy.textContent = "Copiado!";
                btnCopy.classList.add("copied");
                setTimeout(() => {
                    btnCopy.textContent = "Copiar Bloco";
                    btnCopy.classList.remove("copied");
                }, 1400);
            });
        }

        function desenharModoEdicao() {
            card.classList.add("editing");

            card.innerHTML = `
        <div class="card-top">
          <button type="button" class="btn-lock-edit unlocked" title="Salvar alterações">
            <span>🔓</span>
            <span>Salvar</span>
          </button>
        </div>

        <div class="card-body">
          <div>
            <div class="col-label">Código</div>
            <input type="text" class="inline-cod-input" value="${escapeAttr(item.codigo)}" />
          </div>

          <div>
            <div class="col-label">Equip.</div>
            <select class="inline-equip-select">
              <option value="GW" ${item.equipamento === "GW" ? "selected" : ""}>GW</option>
              <option value="KHOMP" ${item.equipamento === "KHOMP" ? "selected" : ""}>KHOMP</option>
            </select>
          </div>

          <div>
            <div class="col-label">Nome da Empresa</div>
            <input type="text" class="inline-input inline-input-empresa" value="${escapeAttr(item.empresa)}" />
          </div>

          <div>
            <div class="col-label">IP / Porta</div>
            <input type="text" class="inline-input inline-input-ip" value="${escapeAttr(item.ip)}" />
          </div>

          <div>
            <div class="col-label">Senha</div>
            <input type="text" class="inline-input inline-input-senha" value="${escapeAttr(item.senha)}" />
          </div>

          <div>
            <button type="button" class="btn-copy" disabled style="opacity: 0.45; cursor: not-allowed;">Copiar Bloco</button>
          </div>
        </div>
      `;

            const inputCod = card.querySelector(".inline-cod-input");
            const selectEquip = card.querySelector(".inline-equip-select");
            const inputEmpresa = card.querySelector(".inline-input-empresa");
            const inputIp = card.querySelector(".inline-input-ip");
            const inputSenha = card.querySelector(".inline-input-senha");

            function salvarEdicaoInline() {
                item.codigo = inputCod.value.trim() || "-";
                item.equipamento = selectEquip.value === "KHOMP" ? "KHOMP" : "GW";
                item.empresa = inputEmpresa.value.trim() || "Sem Nome";
                item.ip = inputIp.value.trim() || "-";
                item.senha = inputSenha.value.trim() || "-";

                salvarStorage();
                desenharModoLeitura();
            }

            card.querySelector(".btn-lock-edit").addEventListener("click", salvarEdicaoInline);

            card.querySelectorAll("input").forEach((inp) => {
                inp.addEventListener("keydown", (e) => {
                    if (e.key === "Enter") {
                        e.preventDefault();
                        salvarEdicaoInline();
                    } else if (e.key === "Escape") {
                        desenharModoLeitura();
                    }
                });
            });

            inputCod.focus();
            inputCod.select();
        }

        desenharModoLeitura();
        cardsContainer.appendChild(card);
    });
}

function escapeHtml(str) {
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function escapeAttr(str) {
    return escapeHtml(str).replace(/'/g, "&#39;");
}

fileDrop.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    state.arquivoPendente = file || null;
    fileNameLabel.textContent = file ? file.name : "Nenhum arquivo selecionado";
    if (file && !nomeBaseInput.value.trim()) {
        nomeBaseInput.value = file.name.replace(/\.[^/.]+$/, "");
    }
});

btnSubirPlanilha.addEventListener("click", async () => {
    if (!state.arquivoPendente) {
        alert("Selecione uma planilha (.xlsx ou .csv) primeiro.");
        return;
    }

    const nomeBase =
        nomeBaseInput.value.trim() ||
        state.arquivoPendente.name.replace(/\.[^/.]+$/, "") ||
        "Base_Importada";

    try {
        const registros = await processarArquivo(state.arquivoPendente);
        if (registros.length === 0) {
            alert("Nenhum registro válido de GW ou KHOMP foi encontrado na planilha.");
            return;
        }

        state.bases[nomeBase] = registros;
        state.baseAtiva = nomeBase;
        salvarStorage();
        renderSelectBases();
        renderCards();

        fileInput.value = "";
        state.arquivoPendente = null;
        fileNameLabel.textContent = "Nenhum arquivo selecionado";
        nomeBaseInput.value = "";

        alert(`Base "${nomeBase}" importada com ${registros.length} registro(s)!`);
    } catch (err) {
        console.error(err);
        alert("Erro ao ler a planilha. Verifique se o formato é válido.");
    }
});

baseSelect.addEventListener("change", (e) => {
    state.baseAtiva = e.target.value;
    salvarStorage();
    renderCards();
});

btnExcluirBase.addEventListener("click", () => {
    if (!state.baseAtiva || !state.bases[state.baseAtiva]) return;
    const nome = state.baseAtiva;
    if (!confirm(`Deseja realmente excluir a base "${nome}"?`)) return;

    delete state.bases[nome];
    const restantes = Object.keys(state.bases);
    state.baseAtiva = restantes[0] || "";
    salvarStorage();
    renderSelectBases();
    renderCards();
});

searchInput.addEventListener("input", (e) => {
    state.busca = e.target.value;
    renderCards();
});

document.querySelectorAll(".pill").forEach((pill) => {
    pill.addEventListener("click", () => {
        document.querySelectorAll(".pill").forEach((p) => p.classList.remove("active"));
        pill.classList.add("active");
        state.filtroEquip = pill.dataset.filter;
        renderCards();
    });
});

btnLimparFiltros.addEventListener("click", () => {
    state.busca = "";
    state.filtroEquip = "TODOS";
    searchInput.value = "";
    document.querySelectorAll(".pill").forEach((p) => {
        p.classList.toggle("active", p.dataset.filter === "TODOS");
    });
    renderCards();
});

carregarStorage();
renderSelectBases();
renderCards();