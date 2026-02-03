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

    // --- 1. LIMPEZA PREVENTIVA DE QUANTIDADES ---
    // Garante que pedidos do colaborador anterior não fiquem sobrando
    for (let i = 0; i < 20; i++) {
        const elQtd = document.getElementById('qtd-' + i);
        if (elQtd) elQtd.value = "";
    }

    // --- 2. CAPTURA E TRATAMENTO DE DADOS ---
    const nome = (linha['NOME'] || linha['nome'] || "").trim();
    const funcao = linha['FUNÇÃO'] || linha['função'] || "";
    const cpf = linha['CPF'] || linha['cpf'] || "";
    const matri = linha['N° Folha'] || linha['Nº Folha'] || "";
    const ctps = linha['CTPS'] || linha['ctps'] || "";
    const ano = linha['ANO'] || linha['ano'] || "2026";
    const setor = linha['SETOR'] || linha['setor'] || "CARPINTARIA"; // Valor padrão
    
    let adm = linha['DATA DE ADMISSÃO'] || linha['DATA DE ADMIMSSÃO'] || "";
    if(typeof adm === 'number') {
        adm = new Date(Math.round((adm - 25569) * 864e5)).toLocaleDateString('pt-BR');
    }

    const setSafe = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val;
    };

    // --- 3. LÓGICA DE MODO DE IMPRESSÃO ---
    const modo = document.getElementById('modoImpressao').value;

    // Campos que devem ser preenchidos nas cabeçalhos
    const campos = ['nome', 'funcao', 'cpf', 'adm', 'matri', 'ano', 'ctps', 'setor'];
    const valores = [nome, funcao, cpf, adm, matri, ano, ctps, setor];

    campos.forEach((campo, i) => {
        // Preenche sempre a Ficha 1 (Principal)
        setSafe(`c-${campo}`, valores[i]);

        // SÓ preenche a Ficha 2 (Nova) se NÃO for reimpressão
        if (modo !== 'reimpressao') {
            setSafe(`c2-${campo}`, valores[i]);
        } else {
            // Se for reimpressão, limpa os campos da ficha 2
            setSafe(`c2-${campo}`, ""); 
        }
    });

    // --- 4. ASSINATURA E FOTO ---
    const sig = document.getElementById('nome-assinatura');
    if (sig) {
        sig.innerText = (modo === 'reimpressao') ? "" : nome;
    }

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

    // --- 5. CORES POR CARGO ---
    const inputFuncao1 = document.getElementById('c-funcao');
    const inputFuncao2 = document.getElementById('c2-funcao');
    if (typeof aplicarCorCargo === "function") {
        aplicarCorCargo(inputFuncao1, funcao);
        if (modo !== 'reimpressao' && inputFuncao2) {
            aplicarCorCargo(inputFuncao2, funcao);
        }
    }

    // --- 6. DISPARA BUSCA NA NUVEM ---
    // A função buscarDadosNuvem agora cuidará de colocar o "1" apenas nas linhas certas
    buscarDadosNuvem(nome);
}

function aplicarCorCargo(elemento, cargo) {
    if (!elemento) return;
    
    const cargoTexto = cargo.toUpperCase();
    let corFundo = "#2e7d32"; // Verde padrão
    let corTexto = "white";

    if (cargoTexto.includes("AUXILIAR")) {
        corFundo = "yellow";
        corTexto = "black";
    } else if (cargoTexto.includes("CARPINTEIRO")) {
        corFundo = "red";
        corTexto = "white";
    } else if (cargoTexto.includes("ENCARREGADO")) {
        corFundo = "blue";
        corTexto = "white";
    }

    // Aplica a cor na tela
    elemento.style.backgroundColor = corFundo;
    elemento.style.color = corTexto;
    
    // Define a variável para o CSS de impressão usar
    elemento.style.setProperty('--cor-bg-funcao', corFundo);
    elemento.style.setProperty('--cor-txt-funcao', corTexto);
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

    const formatarDataExcel = (valor) => {
        if (!valor) return "";
        if (!isNaN(valor) && typeof valor === 'number') {
            const data = new Date(Math.round((valor - 25569) * 864e5));
            data.setMinutes(data.getMinutes() + data.getTimezoneOffset());
            return data.toLocaleDateString('pt-BR');
        }
        if (typeof valor === 'string' && valor.includes('-')) {
            const partes = valor.split('T')[0].split('-');
            if (partes.length === 3) return `${partes[2]}/${partes[1]}/${partes[0]}`;
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
        const pular = parseInt(document.getElementById('pularLinhas').value) || 0;

        const filtrados = dadosNuvem.filter(linha => {
            const f = linha.funcionario || linha.Funcionario || Object.values(linha)[0];
            const dRaw = linha.dataPedido || linha.DATA || Object.values(linha)[2];
            const dISO = dRaw ? dRaw.toString().split('T')[0] : "";
            return limpar(f) === nomeAlvo && (!dInicio || dISO >= dInicio) && (!dFim || dISO <= dFim);
        });

        // --- CORREÇÃO: Limpa TODAS as colunas, incluindo 'qtd-', antes de preencher ---
        for (let i = 0; i < 20; i++) {
            ['data-', 'desc-', 'fab-', 'ca-', 'val-', 'dev-', 'qtd-'].forEach(p => { 
                const el = document.getElementById(p + i);
                if (el) el.value = ""; 
            });
        }

        if (filtrados.length > 0) {
            filtrados.forEach((reg, index) => {
                const linhaAlvo = index + pular;

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

                    const setV = (p, v) => { 
                        const el = document.getElementById(p + linhaAlvo); 
                        if(el) el.value = v; 
                    };
                    
                    // PREENCHIMENTO DOS DADOS
                    setV('data-', formatarDataExcel(reg.dataPedido || colNuvem[2]));
                    setV('desc-', epiNuvem);
                    setV('dev-', "NÃO");
                    setV('qtd-', "1"); // Preenche apenas nas linhas que contêm dados

                    if (infoLocal) {
                        Object.keys(infoLocal).forEach(key => {
                            const k = limpar(key);
                            let val = infoLocal[key];
                            if (k === "FABRICANTE") setV('fab-', val || "NÃO");
                            if (k === "CA" || k === "C.A") setV('ca-', val || "");
                            if (k.includes("VALIDADE") && (k.includes("C.A") || k.includes("CA"))) {
                                setV('val-', formatarDataExcel(val));
                            }
                        });
                    } else {
                        setV('fab-', "NÃO");
                    }
                }
            });

            // --- GARANTIR O SETOR CARPINTARIA SE ESTIVER VAZIO ---
            const cSetor = document.getElementById('c-setor');
            if (cSetor && !cSetor.value) cSetor.value = "CARPINTARIA";

            corpo.innerHTML = `✅ Ficha preenchida (${filtrados.length} itens)!`;
            setTimeout(() => modal.style.display='none', 1000);
        } else {
            corpo.innerHTML = "Nenhum dado encontrado para este período.";
            setTimeout(() => modal.style.display='none', 2000);
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
    // Limpa a tabela antes de preencher
    for (let i = 0; i < 20; i++) {
        ['dev-', 'data-', 'desc-', 'qtd-', 'fab-', 'ca-', 'val-'].forEach(prefix => {
            const el = document.getElementById(prefix + i);
            if (el) el.value = "";
        });
    }

    if (valor === "nao") return;

    if (valor === "inicial") {
        if (!dadosPlanilha || dadosPlanilha.length === 0) {
            alert("⚠️ Carregue a planilha Excel primeiro!");
            return;
        }

        const dataHoje = new Date().toLocaleDateString('pt-BR');
        
        // FUNÇÃO MELHORADA PARA NÃO PEGAR COLUNA ERRADA
        const getVal = (obj, busca) => {
            const termo = busca.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
            const key = Object.keys(obj).find(k => {
                const nomeCol = k.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
                // Se buscamos CA, precisamos que seja EXATO ou contenha C.A (com pontos)
                if (termo === "CA") {
                    return nomeCol === "CA" || nomeCol === "C.A" || nomeCol === "C.A.";
                }
                return nomeCol.includes(termo);
            });
            return key ? obj[key] : "";
        };

        dadosPlanilha.forEach((linha, i) => {
            if (i < 20) {
                const desc = getVal(linha, "DESCRICAO") || getVal(linha, "EPI") || getVal(linha, "PRODUTO");
                const fab  = getVal(linha, "FABRICANTE");
                const ca   = getVal(linha, "CA"); // Agora a busca é mais rigorosa aqui
                let val    = getVal(linha, "VALIDADE");

                if (desc) {
                    let qtdDefinida = "1";
                    const descUpper = desc.toString().toUpperCase();
                    
                    if (descUpper.includes("CAMISA")) {
                        qtdDefinida = "2";
                    } else if (descUpper.includes("CALÇA")) {
                        qtdDefinida = "3";
                    }

                    if(val && typeof val === 'number') {
                        val = new Date(Math.round((val - 25569) * 864e5)).toLocaleDateString('pt-BR');
                    }

                    const setV = (id, v) => { 
                        const el = document.getElementById(id);
                        if(el) el.value = v; 
                    };

                    setV(`dev-${i}`, "INICIAL");
                    setV(`data-${i}`, dataHoje);
                    setV(`desc-${i}`, desc);
                    setV(`qtd-${i}`, qtdDefinida);
                    setV(`fab-${i}`, fab || "COELHO");
                    setV(`ca-${i}`, ca || ""); // Se continuar trazendo função, verifique o nome da coluna no Excel
                    setV(`val-${i}`, val || "");
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

    // 1. Sincroniza o que foi digitado/preenchido para aparecer no papel
    document.querySelectorAll('input').forEach(input => {
        input.setAttribute('value', input.value);
    });

    // 2. Controla a visibilidade das páginas
    if (modo === 'reimpressao') {
        if (p1) p1.style.display = 'none'; // Esconde a folha de dados (Pag 1)
        if (p2) {
            p2.style.display = 'block';
            p2.classList.add('modo-fantasma'); // Deixa só os textos na Pag 2
        }
    } else {
        // MODO FICHA NOVA: Garante que as duas apareçam
        if (p1) p1.style.display = 'block';
        if (p2) {
            p2.style.display = 'block';
            p2.classList.remove('modo-fantasma');
        }
    }
}

// Chame esta função no seu botão de imprimir no HTML ou adicione o window.print()
function imprimir() {
    ajustarVisibilidadeImpressao();
    window.print();
}


function imprimirFicha() {
    // 1. Pega o nome do colaborador selecionado
    const select = document.getElementById('selectColaborador');
    const nomeColaborador = select.options[select.selectedIndex].text;

    // 2. Define o título da página com o nome dele
    // Isso fará com que o PDF sugira: "FICHA_EPI_JOAO_SILVA.pdf"
    if (nomeColaborador && nomeColaborador !== "-- Colaborador --") {
        document.title = "FICHA_EPI_" + nomeColaborador.toUpperCase().replace(/\s+/g, '_');
    } else {
        document.title = "FICHA_EPI_COELHO_EMPREITEIRA";
    }

    // 3. Aplica as regras de visibilidade que já configuramos
    ajustarVisibilidadeImpressao();

    // 4. Abre a tela de impressão
    window.print();

    // 5. Opcional: Volta o título original após imprimir
    setTimeout(() => {
        document.title = "Ficha EPI - Coelho Empreiteira";
    }, 1000);
}


function otimizarMemoriaImpressao() {
    document.querySelectorAll('input').forEach(input => {
        input.style.outline = "none";
        input.style.boxShadow = "none";
    });
}   
