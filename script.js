let dadosPlanilha = [];
let historicoBuscas = [];
let nomeContagemTotal = {};
let codigosProcessados = {};

// Função para ler a planilha
document.getElementById("input-excel").addEventListener("change", (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        dadosPlanilha = jsonData.slice(1).map(row => ({
            codigo: row[0] || "",
            driver: row[1] || "",
            cliente: row[2] || "" // Adicionando o cliente da coluna C
        }));

        nomeContagemTotal = contarOcorrencias(dadosPlanilha);
        exibirContagem(nomeContagemTotal); // Exibe contagem de nomes
        atualizarHistorico(); // Atualiza o histórico para mostrar todos os nomes
    };

    reader.readAsArrayBuffer(file);
});

// Adiciona um listener para o evento keydown no campo de entrada do código
document.getElementById("codigo").addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
        event.preventDefault();
        buscarPorCodigo();
    }
});

// Função que conta ocorrências de nomes
function contarOcorrencias(planilha) {
    const contagem = {};

    planilha.forEach(item => {
        const nomeDriver = item.driver.trim();

        if (nomeDriver) {
            const nomeBase = nomeDriver.split(" ")[0];

            if (!contagem[nomeBase]) {
                contagem[nomeBase] = new Set();
            }
            contagem[nomeBase].add(nomeDriver);
        }
    });

    for (const nomeBase in contagem) {
        contagem[nomeBase] = contagem[nomeBase].size;
    }

    return contagem;
}

function exibirContagem(contagem) {
    const contagemDiv = document.getElementById("contagem-nomes");
    contagemDiv.innerHTML = "";

    const contenedorFlex = document.createElement("div");
    contenedorFlex.classList.add("contagem-container");

    for (const [nome, quantidade] of Object.entries(contagem)) {
        const p = document.createElement("p");
        p.classList.add("contagem-item");

        p.innerHTML = `<strong>${nome}</strong>: <span class="total">${quantidade}</span>`;
        contenedorFlex.appendChild(p);
    }

    contagemDiv.appendChild(contenedorFlex);
}

// Função para buscar pelo código
function buscarPorCodigo() {
    let codigoInput = document.getElementById("codigo").value.trim();
    codigoInput = codigoInput.replace(/^0+/, ''); // Remove os zeros à esquerda
    const codigoConcat = `${codigoInput}-1`; // Concatena "-1"
    const resultadosDiv = document.getElementById("resultado");
    resultadosDiv.innerHTML = "";

    if (codigoInput === "") {
        resultadosDiv.innerHTML = "Por favor, insira um código para buscar.";
        return;
    }

    const resultado = dadosPlanilha.find(item => item.codigo.toString().replace(/^0+/, '') === codigoConcat);

    if (resultado) {
        const nomeDriver = resultado.driver.trim();
        resultadosDiv.innerHTML = `Código: <strong style="color: green;">${resultado.codigo}</strong> - Driver: <span class="highlight">${resultado.driver}</span>`;

        historicoBuscas.unshift({
            codigo: resultado.codigo,
            driver: resultado.driver,
            cliente: resultado.cliente, // Adicionando o cliente ao histórico
            sucesso: true // Marca como pesquisa bem-sucedida
        });
        atualizarHistorico();

        speakText(`${resultado.driver}`);
        imprimirDriver(resultado.driver);

        const nomeBase = nomeDriver.split(" ")[0];

        if (!codigosProcessados[codigoInput] && nomeContagemTotal[nomeBase] > 0) {
            nomeContagemTotal[nomeBase]--;
            codigosProcessados[codigoInput] = true;
            exibirContagem(nomeContagemTotal);
        }

    } else {
        resultadosDiv.innerHTML = "Nenhum resultado encontrado.";
        // Não adiciona ao histórico em caso de falha na pesquisa
    }

    document.getElementById("codigo").value = "";
    document.getElementById("codigo").focus();
}

function imprimirDriver(driver) {
    const iframe = document.createElement('iframe');
    iframe.style.position = 'absolute';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = 'none';
    document.body.appendChild(iframe);

    const estiloEtiquetadora = `
        <style>
            body {
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
                font-family: "Arial", sans-serif;
                overflow: hidden;
                background-color: white;
            }
            h1 {
                font-size: 30px;
                font-weight: bold;
                text-align: center;
                border: 2px solid black;
                padding: 20px;
                width: 60%;
                margin: 0;
            }
            @media print {
                body {
                    width: 100%;
                    height: 100%;
                }
            }
        </style>
    `;

    const conteudo = `${estiloEtiquetadora}<h1>${driver}</h1>`;
    const doc = iframe.contentWindow.document;
    doc.open();
    doc.write(conteudo);
    doc.close();

    iframe.contentWindow.focus();
    iframe.contentWindow.print();
    iframe.parentNode.removeChild(iframe);
}

function speakText(text) {
    const synth = window.speechSynthesis;
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = 'pt-BR'; 
    utterance.rate = 2.0;
    utterance.pitch = 1.3;

    synth.speak(utterance);
}

function atualizarHistorico() {
    const listaHistorico = document.getElementById("lista-historico");
    listaHistorico.innerHTML = "";

    historicoBuscas.forEach(item => {
        const li = document.createElement("li");
        li.classList.add("resultado-historico");

        const backgroundColor = item.sucesso ? '#d0f0c0' : '#fff';
        const textColor = item.sucesso ? '#000' : '#888';

        li.innerHTML = `
            <div style="background-color: ${backgroundColor}; color: ${textColor}; padding: 10px; border-radius: 5px;">
                Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>
                <button class="btn-ver-cliente" onclick="toggleCliente('${item.codigo}')">Ver Cliente</button>
                <div class="cliente" id="cliente-${item.codigo}" style="display: none;">Cliente: ${item.cliente}</div>
            </div>
        `;
        listaHistorico.appendChild(li);
    });
}

// Função para mostrar/ocultar o cliente
function toggleCliente(codigo) {
    const clienteDiv = document.getElementById(`cliente-${codigo}`);
    if (clienteDiv) {
        clienteDiv.style.display = clienteDiv.style.display === "none" ? "block" : "none";
    }
}

// Função para filtrar histórico
function filtrarHistorico() {
    const filtro = document.getElementById("filtro").value.trim().toLowerCase();
    const listaHistorico = document.getElementById("lista-historico");
    listaHistorico.innerHTML = "";

    if (filtro === "") {
        atualizarHistorico();
        return;
    }

    const resultadosFiltrados = dadosPlanilha.filter(item =>
        item.driver.toLowerCase().includes(filtro) || item.codigo.toString().replace(/^0+/, '').concat('-1').includes(filtro)
    );

    resultadosFiltrados.forEach(item => {
        const li = document.createElement("li");
        li.classList.add("resultado-historico");

        const backgroundColor = historicoBuscas.some(h => h.codigo === item.codigo && h.sucesso) ? '#d0f0c0' : '#f8d7da';
        const textColor = backgroundColor === '#d0f0c0' ? '#000' : '#888';

        li.innerHTML = `
            <div style="background-color: ${backgroundColor}; color: ${textColor}; padding: 10px; border-radius: 5px;">
                Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>
                <button class="btn-ver-cliente" onclick="toggleCliente('${item.codigo}')">Ver Cliente</button>
                <div class="cliente" id="cliente-${item.codigo}" style="display: none;">Cliente: ${item.cliente}</div>
            </div>
        `;
        listaHistorico.appendChild(li);
    });
}
