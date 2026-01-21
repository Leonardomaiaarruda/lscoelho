/**
 * SISTEMA DE FICHA DE EPI - COELHO EMPREITEIRA
 * Versão Otimizada para Performance de Impressão
 */

let dadosPlanilha = [];
let bibliotecaFotos = {};

// 1. INICIALIZAÇÃO SEGURA
document.addEventListener('DOMContentLoaded', () => {
    const campoData = document.getElementById('data-atual');
    if (campoData) campoData.value = new Date().toLocaleDateString('pt-BR');

    // Configuração dos Listeners de Arquivo
    const inputExcel = document.getElementById('inputExcel');
    if (inputExcel) inputExcel.addEventListener('change', carregarExcel);

    const inputPasta = document.getElementById('inputPasta');
    if (inputPasta) inputPasta.addEventListener('change', carregarPastaFotos);
});

// 2. CARREGAMENTO DO EXCEL
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

// 3. CARREGAMENTO DE FOTOS (COM LIMPEZA DE MEMÓRIA)
function carregarPastaFotos(e) {
    // Revoga URLs anteriores para não travar o navegador com lixo de memória
    Object.values(bibliotecaFotos).forEach(url => URL.revokeObjectURL(url));
    bibliotecaFotos = {}; 
    
    for (let f of e.target.files) {
        // Salva com nome limpo (minúsculo e sem espaços nas pontas)
        bibliotecaFotos[f.name.toLowerCase().trim()] = URL.createObjectURL(f);
    }
    console.log("Biblioteca de fotos carregada.");
}

// 4. PREENCHIMENTO AUTOMÁTICO DA FICHA
function preencher() {
    const idx = document.getElementById('selectColaborador').value;
    if(idx === "" || !dadosPlanilha[idx]) return;
    const linha = dadosPlanilha[idx];

    // Mapeamento de campos
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

    // Função interna para evitar erro "null" ao setar valores
    const setSafe = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val;
    };

    // Preenchimento Ficha 1 e 2
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
        let ref = (linha['FOTO'] || linha['foto'] || nome).toString().toLowerCase().trim();
        let url = bibliotecaFotos[ref] || 
                  bibliotecaFotos[ref + ".jpg"] || 
                  bibliotecaFotos[ref + ".png"] || 
                  bibliotecaFotos[ref + ".jpeg"];

        if (url) {
            img.src = url;
            img.style.display = 'block';
            if (placeholder) placeholder.style.display = 'none';
        } else {
            img.style.display = 'none';
            if (placeholder) placeholder.style.display = 'block';
        }
    }
}

// 5. PROCESSAMENTO DE EPIS (INICIAL)
function processarEntregaExcel(selectElement) {
    if (!selectElement || selectElement.value !== "inicial") return;

    if (dadosPlanilha.length === 0) {
        alert("⚠️ Carregue a planilha Excel primeiro!");
        selectElement.value = "nao";
        return;
    }

    const dataHoje = new Date().toLocaleDateString('pt-BR');
    const totalLinhasFicha = 20;

    const formatarData = (v) => {
        if (!v) return "";
        if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5)).toLocaleDateString('pt-BR');
        return v;
    };

    dadosPlanilha.forEach((linha, i) => {
        if (i < totalLinhasFicha) {
            const desc = linha['DESCRIÇÃO DO EPI'] || linha['DESCRIÇÃO'] || linha['ITEM'] || "";
            if (desc) {
                if (i > 0) {
                    const dev = document.getElementById(`dev-${i}`);
                    if (dev) dev.value = "INICIAL";
                }
                const campoData = document.getElementById(`data-${i}`);
                if (campoData) campoData.value = formatarData(linha['DATA DE ENTREGA'] || linha['DATA']) || dataHoje;
                
                const campoDesc = document.getElementById(`desc-${i}`);
                if (campoDesc) campoDesc.value = desc;

                const campoQtd = document.getElementById(`qtd-${i}`);
                if (campoQtd) campoQtd.value = linha['QTD'] || linha['QUANTIDADE'] || "1";

                const campoFab = document.getElementById(`fab-${i}`);
                if (campoFab) campoFab.value = linha['FABRICANTE'] || "";

                // Busca CA e Validade dinamicamente
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

// 6. ACELERAÇÃO DE IMPRESSÃO (Otimiza a renderização do PDF)
window.onbeforeprint = () => {
    const inputs = document.querySelectorAll('input, select');
    inputs.forEach(el => {
        if (el.tagName === 'INPUT') {
            el.setAttribute('value', el.value);
        }
        if (el.tagName === 'SELECT') {
            const selectedOption = el.options[el.selectedIndex];
            if (selectedOption) {
                for (let opt of el.options) opt.removeAttribute('selected');
                selectedOption.setAttribute('selected', 'selected');
            }
        }
    });
};
