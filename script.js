/**
 * SISTEMA DE FICHA DE EPI - COELHO EMPREITEIRA
 * Fotos via Google Drive | Dados via Excel Local
 */

let dadosPlanilha = [];

// 1. INICIALIZAÇÃO
document.addEventListener('DOMContentLoaded', () => {
    const campoData = document.getElementById('data-atual');
    if (campoData) campoData.value = new Date().toLocaleDateString('pt-BR');

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
        // Link direto via thumbnail (mais estável para navegadores)
        return `https://lh3.googleusercontent.com/d/${matches[1]}`;
    }
    return link; 
}

// 3. CARREGAMENTO DO EXCEL
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

// 4. PREENCHIMENTO AUTOMÁTICO (DADOS PESSOAIS)
function preencher() {
    const idx = document.getElementById('selectColaborador').value;
    if(idx === "" || !dadosPlanilha[idx]) return;
    const linha = dadosPlanilha[idx];

    const nome = linha['NOME'] || linha['nome'] || "";
    const funcao = linha['FUNÇÃO'] || linha['função'] || "";
    const cpf = linha['CPF'] || linha['cpf'] || "";
    const matri = linha['N° Folha'] || linha['Nº Folha'] || "";
    const ctps = linha['CTPS'] || linha['ctps'] || "";
    const ano = linha['ANO'] || linha['ano'] || "2026";
    
    let adm = linha['DATA DE ADMISSÃO'] || linha['DATA DE ADMIMSSÃO'] || "";
    if(typeof adm === 'number') {
        adm = new Date(Math.round((adm - 25569) * 864e5)).toLocaleDateString('pt-BR');
    }

    const setSafe = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val;
    };

    const campos = ['nome', 'funcao', 'cpf', 'adm', 'matri', 'ano', 'ctps'];
    const valores = [nome, funcao, cpf, adm, matri, ano, ctps];

    campos.forEach((campo, i) => {
        setSafe(`c-${campo}`, valores[i]);
        setSafe(`c2-${campo}`, valores[i]);
    });

    const sig = document.getElementById('nome-assinatura');
    if (sig) sig.innerText = nome;

    // Lógica da Foto
    const img = document.getElementById('img-colab');
    const placeholder = document.getElementById('placeholder-foto');
    if (img) {
        const linkBruto = linha['FOTO'] || linha['foto'] || "";
        const urlDireta = formatarLinkDrive(linkBruto);
        if (urlDireta) {
            img.src = urlDireta;
            img.style.display = 'block';
            if (placeholder) placeholder.style.display = 'none';
            img.onerror = () => {
                img.style.display = 'none';
                if (placeholder) {
                    placeholder.style.display = 'block';
                    placeholder.innerText = "ERRO DRIVE";
                }
            };
        } else {
            img.style.display = 'none';
            if (placeholder) {
                placeholder.style.display = 'block';
                placeholder.innerText = "SEM FOTO";
            }
        }
    }
}

// 5. PROCESSAMENTO DE EPIS (CONTROLADO PELO MENU SUPERIOR)
function processarEntregaExcel(valor) {
    // 1. LÓGICA DE LIMPEZA (Se valor for "nao")
    if (valor === "nao") {
        for (let i = 0; i < 20; i++) {
            const campos = [`dev-${i}`, `data-${i}`, `desc-${i}`, `qtd-${i}`, `fab-${i}`, `ca-${i}`, `val-${i}`];
            campos.forEach(id => {
                const el = document.getElementById(id);
                if (el) el.value = "";
            });
        }
        return;
    }

    // 2. LÓGICA DE PREENCHIMENTO (Se valor for "inicial")
    if (valor === "inicial") {
        if (dadosPlanilha.length === 0) {
            alert("⚠️ Carregue a planilha Excel primeiro!");
            return;
        }

        const dataHoje = new Date().toLocaleDateString('pt-BR');
        const formatarData = (v) => {
            if (!v) return "";
            if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5)).toLocaleDateString('pt-BR');
            return v;
        };

        dadosPlanilha.forEach((linha, i) => {
            if (i < 20) {
                const desc = linha['DESCRIÇÃO DO EPI'] || linha['DESCRIÇÃO'] || linha['ITEM'] || "";
                if (desc) {
                    // Preenche "INICIAL" na coluna devolução
                    const dev = document.getElementById(`dev-${i}`);
                    if (dev) dev.value = "INICIAL";

                    const campoData = document.getElementById(`data-${i}`);
                    if (campoData) campoData.value = formatarData(linha['DATA DE ENTREGA'] || linha['DATA']) || dataHoje;
                    
                    const campoDesc = document.getElementById(`desc-${i}`);
                    if (campoDesc) campoDesc.value = desc;

                    const campoQtd = document.getElementById(`qtd-${i}`);
                    if (campoQtd) campoQtd.value = linha['QTD'] || linha['QUANTIDADE'] || "1";

                    const campoFab = document.getElementById(`fab-${i}`);
                    if (campoFab) campoFab.value = linha['FABRICANTE'] || "";

                    const colCA = Object.keys(linha).find(k => k.toUpperCase().includes('C.A'));
                    if (colCA) {
                        const elCA = document.getElementById(`ca-${i}`);
                        if (elCA) elCA.value = linha[colCA];
                    }

                    const colVal = Object.keys(linha).find(k => k.toUpperCase().includes('VAL'));
                    if (colVal) {
                        const elVal = document.getElementById(`val-${i}`);
                        if (elVal) elVal.value = formatarData(linha[colVal]);
                    }
                }
            }
        });
    }
}

// 6. SINCRONIZAÇÃO DO MENU SUPERIOR
function sincronizarMobile(elMobile) {
    // Agora enviamos o valor direto para a função de processamento
    processarEntregaExcel(elMobile.value);
}

// 7. AJUSTE DE IMPRESSÃO
window.onbeforeprint = () => {
    const inputs = document.querySelectorAll('input, select');
    inputs.forEach(el => {
        if (el.tagName === 'INPUT') el.setAttribute('value', el.value);
        if (el.tagName === 'SELECT') {
            const selectedOption = el.options[el.selectedIndex];
            if (selectedOption) {
                for (let opt of el.options) opt.removeAttribute('selected');
                selectedOption.setAttribute('selected', 'selected');
            }
        }
    });
};
