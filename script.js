/**
 * SISTEMA DE FICHA DE EPI - COELHO EMPREITEIRA
 * Fotos via Google Drive | Dados via Excel Local | Consulta Nuvem via Google Apps Script
 */

let dadosPlanilha = [];
// URL do seu Google Apps Script
const URL_NUVEM = "https://script.google.com/macros/s/AKfycbyZPyhDd70Ez-KbJBBTl07Vffpf6Vl2Qexi00Qh1BJdIFbHU7aq50ONE74GEVpeqMZIZg/exec";

// 1. INICIALIZAÇÃO
document.addEventListener('DOMContentLoaded', () => {
    const inputExcel = document.getElementById('inputExcel');
    if (inputExcel) inputExcel.addEventListener('change', carregarExcel);
});

// 2. FUNÇÃO PARA FORMATAR LINK DO DRIVE (CORRIGIDA)
function formatarLinkDrive(link) {
    if (!link) return "";
    link = link.toString().trim();
    const regExp = /(?:id=|\/d\/)([\w-]+)/;
    const matches = link.match(regExp);
    if (matches && matches[1]) {
        // Corrigido: Adicionado o $ para a variável matches[1]
        return `https://lh3.googleusercontent.com/u/0/d/${matches[1]}`;
    }
    return link; 
}

// 3. CARREGAMENTO DO EXCEL LOCAL
function carregarExcel(e) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const workbook = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
        dadosPlanilha = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        const sel = document.getElementById('selectColaborador');
        if (sel) {
            sel.innerHTML = '<option value="">Selecione...</option>';
            dadosPlanilha.forEach((d, i) => {
                const nomeColab = d.NOME || d.nome || "Sem Nome";
                sel.innerHTML += `<option value="${i}">${nomeColab}</option>`;
            });
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
}

// 4. PREENCHIMENTO AUTOMÁTICO E BUSCA NA NUVEM
async function preencher() {
    const select = document.getElementById('selectColaborador');
    if (select.disabled) return; 

    const idx = select.value;
    if (idx === "" || !dadosPlanilha[idx]) return;

    const linha = dadosPlanilha[idx];

    // 3. Captura e tratamento de dados básicos
    const nome = (linha['NOME'] || linha['nome'] || "").trim();
    const funcao = linha['FUNÇÃO'] || linha['função'] || "";
    const cpf = linha['CPF'] || linha['cpf'] || "";
    const matri = linha['N° Folha'] || linha['Nº Folha'] || "";
    const ctps = linha['CTPS'] || linha['ctps'] || "";
    const ano = linha['ANO'] || linha['ano'] || "2026";
    const setor = linha['SETOR'] || linha['setor'] || ""; // Captura o setor
    
    let adm = linha['DATA DE ADMISSÃO'] || linha['DATA DE ADMIMSSÃO'] || "";
    if(typeof adm === 'number') {
        adm = new Date(Math.round((adm - 25569) * 864e5)).toLocaleDateString('pt-BR');
    }

    const setSafe = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val;
    };

    // --- LOGICA DE MODO DE IMPRESSÃO ---
    const modo = document.getElementById('modoImpressao').value;

    // Adicionado 'setor' na lista para limpeza
    const campos = ['nome', 'funcao', 'cpf', 'adm', 'matri', 'ano', 'ctps', 'setor'];
    const valores = [nome, funcao, cpf, adm, matri, ano, ctps, setor];

    campos.forEach((campo, i) => {
        // Preenche sempre a Ficha 1
        setSafe(`c-${campo}`, valores[i]);

        // SÓ preenche a Ficha 2 se NÃO for reimpressão
        if (modo !== 'reimpressao') {
            setSafe(`c2-${campo}`, valores[i]);
        } else {
            setSafe(`c2-${campo}`, ""); // Limpa Nome, CPF e SETOR se for reimpressão
        }
    });

    // 6. Atualiza o nome na área de assinatura (Só mostra se não for reimpressão)
    const sig = document.getElementById('nome-assinatura');
    if (sig) {
        sig.innerText = (modo === 'reimpressao') ? "" : nome;
    }

    // 7. Lógica da Foto
    const img = document.getElementById('img-colab');
    const placeholder = document.getElementById('placeholder-foto');
    if (img) {
        const linkBruto = linha['FOTO'] || linha['foto'] || "";
        const urlDireta = formatarLinkDrive(linkBruto);
        if (urlDireta) {
            img.src = urlDireta;
            img.style.display = (modo === 'reimpressao') ? 'none' : 'block';
            if (placeholder) placeholder.style.display = 'none';
        } else {
            img.src = "";
            img.style.display = 'none';
            if (placeholder) {
                placeholder.style.display = (modo === 'reimpressao') ? 'none' : 'block';
                placeholder.innerText = "SEM FOTO";
            }
        }
    }

    // 8. Cores por Cargo
    const inputFuncao1 = document.getElementById('c-funcao');
    const inputFuncao2 = document.getElementById('c2-funcao');
    if (typeof aplicarCorCargo === "function") {
        aplicarCorCargo(inputFuncao1, funcao);
        if (modo !== 'reimpressao' && inputFuncao2) {
            aplicarCorCargo(inputFuncao2, funcao);
        }
    }

    // 9. Dispara busca na Nuvem
    buscarDadosNuvem(nome);
} 

function aplicarCorCargo(elemento, cargo) {
    if (!elemento) return;
    elemento.style.backgroundColor = "";
    elemento.style.color = "black";
    const cargoTexto = cargo.toUpperCase();
    if (cargoTexto.includes("AUXILIAR")) {
        elemento.style.backgroundColor = "yellow";
    } else if (cargoTexto.includes("CARPINTEIRO")) {
        elemento.style.backgroundColor = "red";
        elemento.style.color = "white";
    } else if (cargoTexto.includes("ENCARREGADO")) {
        elemento.style.backgroundColor = "blue";
        elemento.style.color = "white";
    }
}

async function buscarDadosNuvem(nomeBusca) {
    let modal = document.getElementById('modalInfo');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'modalInfo';
        modal.style = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); z-index:9999; display:none; justify-content:center; align-items:center;";
        modal.innerHTML = `<div style="background:white; padding:20px; border-radius:10px; max-width:450px; width:90%; color:black; position:relative;"><div id="modal-body"></div></div>`;
        document.body.appendChild(modal);
    }
    
    const corpo = document.getElementById('modal-body');
    modal.style.display = 'flex';
    corpo.innerHTML = "⌛ Sincronizando dados e calculando linhas...";

    // Função interna para converter número do Excel em Data BR
    const formatarDataExcel = (valor) => {
        if (!valor) return "";

        if (!isNaN(valor) && typeof valor === 'number') {
            const data = new Date(Math.round((valor - 25569) * 864e5));
            data.setMinutes(data.getMinutes() + data.getTimezoneOffset());
            return data.toLocaleDateString('pt-BR');
        }

        if (typeof valor === 'string' && valor.includes('-')) {
            const partes = valor.split('T')[0].split('-');
            if (partes.length === 3) {
                return `${partes[2]}/${partes[1]}/${partes[0]}`;
            }
        }
        return valor;
    };

    try {
        const response = await fetch(URL_NUVEM);
        const dadosNuvem = await response.json();
        const limpar = (t) => t ? t.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim() : "";
        const nomeAlvo = limpar(nomeBusca);
        const dInicio = document.getElementById('dataInicio').value;
        const dFim = document.getElementById('dataFim').value;
        
        // CAPTURA O VALOR DE LINHAS A PULAR
        const pular = parseInt(document.getElementById('pularLinhas').value) || 0;

        const filtrados = dadosNuvem.filter(linha => {
            const f = linha.funcionario || linha.Funcionario || Object.values(linha)[0];
            const dRaw = linha.dataPedido || linha.DATA || Object.values(linha)[2];
            const dISO = dRaw ? dRaw.toString().split('T')[0] : "";
            return limpar(f) === nomeAlvo && (!dInicio || dISO >= dInicio) && (!dFim || dISO <= dFim);
        });

        // Limpa TODAS as 20 linhas da ficha antes de começar para evitar sobreposição visual
        for (let i = 0; i < 20; i++) {
            ['data-', 'desc-', 'fab-', 'ca-', 'val-', 'dev-'].forEach(p => { 
                const el = document.getElementById(p + i);
                if (el) el.value = ""; 
            });
        }

        if (filtrados.length > 0) {
            filtrados.forEach((reg, index) => {
                // CALCULA A LINHA ALVO (Index atual + Linhas puladas)
                const linhaAlvo = index + pular;

                // Só preenche se ainda houver espaço nas 20 linhas da ficha
                if (linhaAlvo < 20) {
                    const colNuvem = Object.values(reg);
                    const epiNuvem = (reg.epi || colNuvem[1] || "").toString().toUpperCase().trim();
                    const epiBusca = limpar(epiNuvem);

                    const infoLocal = dadosPlanilha.find(item => {
                        return Object.keys(item).some(key => {
                            const k = limpar(key);
                            return (k.includes("EPI") || k.includes("PRODUTO") || k.includes("DESC")) && limpar(item[key]) === epiBusca;
                        });
                    });

                    // Função auxiliar ajustada para usar a 'linhaAlvo'
                    const setV = (p, v) => { 
                        const el = document.getElementById(p + linhaAlvo); 
                        if(el) el.value = v; 
                    };
                    
                    setV('data-', formatarDataExcel(reg.dataPedido || colNuvem[2]));
                    setV('desc-', epiNuvem);
                    setV('dev-', "NÃO");

                    if (infoLocal) {
                        Object.keys(infoLocal).forEach(key => {
                            const k = limpar(key);
                            let val = infoLocal[key];
                            
                            if (k === "FABRICANTE") setV('fab-', val || "NÃO");
                            if (k === "CA" || k === "C.A") setV('ca-', val || "");
                            if (k.includes("VALIDADE") && k.includes("C.A")) {
                                setV('val-', formatarDataExcel(val));
                            }
                        });
                    } else {
                        setV('fab-', "NÃO");
                    }
                }
            });
            corpo.innerHTML = `✅ Ficha preenchida a partir da linha ${pular + 1}!`;
            setTimeout(() => modal.style.display='none', 1000);
        } else {
            corpo.innerHTML = "Nenhum dado encontrado.";
        }
    } catch (e) { 
        corpo.innerHTML = "Erro na sincronização."; 
        console.error(e);
    }
}

// 6. FUNÇÃO PARA FECHAR O MODAL (Conserta o erro de ReferenceError)
function fecharModal() {
    const modal = document.getElementById('modalInfo');
    if (modal) modal.style.display = 'none';
}

// 7. PROCESSAMENTO DE EPIS (LÓGICA LOCAL)
function processarEntregaExcel(valor) {
    if (valor === "nao") {
        for (let i = 0; i < 20; i++) {
            ['dev-', 'data-', 'desc-', 'qtd-', 'fab-', 'ca-', 'val-'].forEach(prefix => {
                const el = document.getElementById(prefix + i);
                if (el) el.value = "";
            });
        }
        return;
    }

    if (valor === "inicial") {
        if (dadosPlanilha.length === 0) {
            alert("⚠️ Carregue a planilha Excel primeiro!");
            return;
        }
        const dataHoje = new Date().toLocaleDateString('pt-BR');
        dadosPlanilha.forEach((linha, i) => {
            if (i < 20) {
                const desc = linha['DESCRIÇÃO DO EPI'] || linha['DESCRIÇÃO'] || "";
                if (desc) {
                    const setVal = (id, val) => { if(document.getElementById(id)) document.getElementById(id).value = val; };
                    setVal(`dev-${i}`, "INICIAL");
                    setVal(`data-${i}`, dataHoje);
                    setVal(`desc-${i}`, desc);
                    setVal(`qtd-${i}`, linha['QTD'] || "1");
                    setVal(`ca-${i}`, linha['C.A'] || "");
                }
            }
        });
    }
}

function sincronizarMobile(elMobile) {
    processarEntregaExcel(elMobile.value);
}

// 8. IMPRESSÃO E OTIMIZAÇÃO
window.onbeforeprint = () => {
    document.querySelectorAll('input, select').forEach(el => {
        if (el.tagName === 'INPUT') el.setAttribute('value', el.value);
        if (el.tagName === 'SELECT') {
            const selected = el.options[el.selectedIndex];
            if (selected) {
                Array.from(el.options).forEach(o => o.removeAttribute('selected'));
                selected.setAttribute('selected', 'selected');
            }
        }
    });
};

// Esta função é o novo "porteiro" do modal
function verificarEBuscar() {
    // Pega o nome que o Excel já escreveu no campo da ficha
    const nomeNaFicha = document.getElementById('c-nome').value;
    const dataFim = document.getElementById('dataFim').value;

    // Se o usuário tentar colocar data sem ter escolhido nome no select
    if (!nomeNaFicha) {
        alert("Por favor, selecione primeiro um colaborador no menu.");
        document.getElementById('dataFim').value = ""; 
        return;
    }

    // O Modal SÓ abre se a data final for preenchida
    if (dataFim) {
        buscarDadosNuvem(nomeNaFicha);
    }
}

function validarDatasParaLiberar() {
    const dataInicio = document.getElementById('dataInicio').value;
    const dataFim = document.getElementById('dataFim').value;
    const selectNome = document.getElementById('selectColaborador');

    // Se ambas as datas estiverem preenchidas
    if (dataInicio !== "" && dataFim !== "") {
        selectNome.disabled = false;
        selectNome.style.border = "2px solid #2e7d32"; // Fica verde para indicar que liberou
        selectNome.options[0].text = "-- Colaborador --";
    } else {
        selectNome.disabled = true;
        selectNome.style.border = "1px solid #ccc";
        selectNome.options[0].text = "Selecione as datas primeiro...";
        selectNome.value = ""; // Reseta a seleção se o usuário apagar uma data
    }
}


function ajustarVisibilidadeImpressao() {
    const modo = document.getElementById('modoImpressao').value;
    const p1 = document.getElementById('pagina-1');
    const p2 = document.getElementById('pagina-tabela');

    if (modo === 'reimpressao') {
        if (p1) p1.style.display = 'none';
        if (p2) p2.classList.add('modo-fantasma');
    } else {
        if (p1) p1.style.display = 'block';
        if (p2) p2.classList.remove('modo-fantasma');
    }

    // Chama a função preencher para limpar os values (Setor, Nome, etc)
    preencher(); 
}

function otimizarMemoriaImpressao() {
    document.querySelectorAll('input').forEach(input => {
        input.style.outline = "none";
        input.style.boxShadow = "none";
    });
}   
