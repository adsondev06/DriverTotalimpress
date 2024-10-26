// let dadosPlanilha = [];
// let historicoBuscas = []; 
// let nomeContagemTotal = {}; 
// let nomesBuscados = {};
// let nomesProcessados = new Set();
// let codigosProcessados = {};

// // Função para ler a planilha
// document.getElementById("input-excel").addEventListener("change", (event) => {
//     const file = event.target.files[0];
//     const reader = new FileReader();

//     reader.onload = (e) => {
//         const data = new Uint8Array(e.target.result);
//         const workbook = XLSX.read(data, { type: 'array' });

//         const worksheet = workbook.Sheets[workbook.SheetNames[0]];
//         const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//         dadosPlanilha = jsonData.slice(1).map(row => ({
//             codigo: row[0], // Coluna A
//             driver: row[1]  // Coluna B
//         }));

//         nomeContagemTotal = contarOcorrencias(dadosPlanilha);

//         // Exibir os resultados na tela
//         exibirContagem(nomeContagemTotal);

//         console.log("Dados importados da planilha:", dadosPlanilha); // Para depuração
//     };

//     reader.readAsArrayBuffer(file);
// });

// // Adiciona um listener para o evento keydown no campo de entrada do código
// document.getElementById("codigo").addEventListener("keydown", function(event) {
//     if (event.key === "Enter") {
//         event.preventDefault(); // Evita o comportamento padrão de envio do formulário
//         buscarPorCodigo(); // Chama a função de busca
//     }
// });


// function contarOcorrencias(planilha) {
//     const contagem = {};

//     planilha.forEach(item => {
//         const nomeDriver = item.driver.trim(); 

//         if (nomeDriver) {
//             const nomeBase = nomeDriver.split(" ")[0];


//             if (!contagem[nomeBase]) {
//                 contagem[nomeBase] = new Set(); 
//             }
//             contagem[nomeBase].add(nomeDriver);
//         }
//     });


//     for (const nomeBase in contagem) {
//         contagem[nomeBase] = contagem[nomeBase].size; 
//     }

//     return contagem;
// }

// function exibirContagem(contagem) {
//     const contagemDiv = document.getElementById("contagem-nomes");
//     contagemDiv.innerHTML = ""; // Limpa resultados anteriores

//     // Cria um contêiner flexível para os itens
//     const contenedorFlex = document.createElement("div");
//     contenedorFlex.classList.add("contagem-container");

//     // Exibe todos os nomes e suas contagens
//     for (const [nome, quantidade] of Object.entries(contagem)) {
//         const p = document.createElement("p");
//         p.classList.add("contagem-item"); // Adiciona uma classe para o estilo

//         p.innerHTML = `<strong>${nome}</strong>: <span class="total">${quantidade}</span>`; 
//         contenedorFlex.appendChild(p);
//     }

//     contagemDiv.appendChild(contenedorFlex); 
// }

// // Função para buscar pelo código
// function buscarPorCodigo() {
//     const codigoInput = document.getElementById("codigo").value.trim();
//     const resultadosDiv = document.getElementById("resultado");
//     resultadosDiv.innerHTML = ""; // Limpa resultados anteriores

//     if (codigoInput === "") {
//         resultadosDiv.innerHTML = "Por favor, insira um código para buscar.";
//         return;
//     }

//     // Busca pelo código
//     const resultado = dadosPlanilha.find(item => item.codigo.toString() === codigoInput);

//     if (resultado) {
//         const nomeDriver = resultado.driver.trim();
//         resultadosDiv.innerHTML = `Código: <strong>${resultado.codigo}</strong> - Driver: <span class="highlight">${resultado.driver}</span>`;

//         historicoBuscas.unshift({
//             codigo: resultado.codigo,
//             driver: resultado.driver
//         });
//         atualizarHistorico();

//         speakText(`${resultado.driver}`);

//         // Chama a função de impressão
//         imprimirDriver(resultado.driver);

//         // Pega o nome base do driver
//         const nomeBase = nomeDriver.split(" ")[0]; 

//         if (!codigosProcessados[nomeBase]) {
//             codigosProcessados[nomeBase] = new Set(); 
//         }

//         // Se o código não tiver sido processado, subtrai 1 da contagem
//         if (!codigosProcessados[nomeBase].has(resultado.codigo) && nomeContagemTotal[nomeBase] > 0) {
//             nomeContagemTotal[nomeBase]--; // Subtrai 1 do total
//             codigosProcessados[nomeBase].add(resultado.codigo); // Marca o código como processado
//             exibirContagem(nomeContagemTotal); // Atualiza a exibição da contagem
//         }

//     } else {
//         resultadosDiv.innerHTML = "Nenhum resultado encontrado.";
//     }

//     // Limpa o campo de busca e mantém o foco
//     document.getElementById("codigo").value = "";
//     document.getElementById("codigo").focus();
// }

// function imprimirDriver(driver) {
//     // Criar um iframe para impressão
//     const iframe = document.createElement('iframe');
//     iframe.style.position = 'absolute';
//     iframe.style.width = '0';
//     iframe.style.height = '0';
//     iframe.style.border = 'none';
//     document.body.appendChild(iframe);

//     // Criar o conteúdo do iframe
//     const estiloEtiquetadora = `
//         <style>
//             body {
//                 display: flex;
//                 justify-content: center;
//                 align-items: center;
//                 height: 100vh;
//                 margin: 0;
//                 font-family: "Arial", sans-serif;
//                 overflow: hidden; /* Impede rolagem */
//                 background-color: white; /* Garante fundo branco */
//             }
//             h1 {
//                 font-size: 30px; /* Tamanho grande para a etiqueta */
//                 font-weight: bold;
//                 text-align: center;
//                 border: 2px solid black; /* Simula borda de etiqueta */
//                 padding: 20px;
//                 width: 60%; /* Controla o tamanho da etiqueta */
//                 margin: 0; /* Remove margem para centrar */
//             }
//             @media print {
//                 body {
//                     width: 100%; /* Controla a largura da impressão */
//                     height: 100%; /* Controla a altura da impressão */
//                 }
//             }
//         </style>
//     `;

//     const conteudo = `${estiloEtiquetadora}<h1>${driver}</h1>`;
//     const doc = iframe.contentWindow.document;
//     doc.open();
//     doc.write(conteudo);
//     doc.close();

//     // Chama a impressão
//     iframe.contentWindow.focus();
//     iframe.contentWindow.print();

//     // Remove o iframe após a impressão
//     iframe.parentNode.removeChild(iframe);
// }

// function speakText(text) {
//     const synth = window.speechSynthesis;
//     const utterance = new SpeechSynthesisUtterance(text);
//     utterance.lang = 'pt-BR'; 
//     utterance.rate = 2.0;
//     utterance.pitch = 1.3; 

//     // Fala o texto
//     synth.speak(utterance);
// }

// function atualizarHistorico() {
//     const listaHistorico = document.getElementById("lista-historico");
//     listaHistorico.innerHTML = "";

//     historicoBuscas.forEach(item => {
//         const li = document.createElement("li");
//         li.innerHTML = `Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>`;
//         listaHistorico.appendChild(li);
//     });
// }

// function filtrarHistorico() {
//     const filtro = document.getElementById("filtro").value.toLowerCase();
//     const listaHistorico = document.getElementById("lista-historico");
//     listaHistorico.innerHTML = "";

//     const historicoFiltrado = historicoBuscas.filter(item => {
//         return item.codigo.toString().toLowerCase().includes(filtro) ||
//                item.driver.toLowerCase().includes(filtro);
//     });

//     // Ordena os resultados
//     historicoFiltrado.sort((a, b) => {
//         const codigoA = a.codigo.toString().toLowerCase();
//         const codigoB = b.codigo.toString().toLowerCase();
//         return codigoA.localeCompare(codigoB);
//     });

//     // Adiciona os itens filtrados à lista
//     historicoFiltrado.forEach(item => {
//         const li = document.createElement("li");
//         li.innerHTML = `Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>`;
//         listaHistorico.appendChild(li);
//     });
// }

// function exportarHistorico() {
//     const workbook = XLSX.utils.book_new();
//     const worksheetData = [
//         ["Código", "Driver"],
//         ...historicoBuscas.map(item => [item.codigo, item.driver])
//     ];
//     const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
//     XLSX.utils.book_append_sheet(workbook, worksheet, "Historico");

//     // Gera o arquivo XLSX
//     XLSX.writeFile(workbook, "historico.xlsx");
// }

// // Botão para reiniciar o histórico de buscas
// document.getElementById("reiniciar-historico").addEventListener("click", function() {
//     historicoBuscas = [];
//     atualizarHistorico();
// });

// // Mostra um alerta de confirmação ao sair
// window.onbeforeunload = function() {
//     return "Tem certeza que deseja sair?"; 
// };




let dadosPlanilha = [];
let historicoBuscas = [];
let nomeContagemTotal = {};
let nomesProcessados = new Set(); // Para rastrear quais nomes já tiveram a contagem reduzida
let codigosProcessados = {}; // Para garantir que cada código único seja processado apenas uma vez

// Função para ler a planilha
document.getElementById("input-excel").addEventListener("change", (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Verificação para evitar `undefined`
        dadosPlanilha = jsonData.slice(1).map(row => ({
            codigo: row[0] || "", // Coluna A
            driver: row[1] || ""  // Coluna B
        }));

        nomeContagemTotal = contarOcorrencias(dadosPlanilha);

        // Exibir os resultados na tela
        exibirContagem(nomeContagemTotal);

        console.log("Dados importados da planilha:", dadosPlanilha); // Para depuração
    };

    reader.readAsArrayBuffer(file);
});

// Adiciona um listener para o evento keydown no campo de entrada do código
document.getElementById("codigo").addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
        event.preventDefault(); // Evita o comportamento padrão de envio do formulário
        buscarPorCodigo(); // Chama a função de busca
    }
});

// Função que conta ocorrências de nomes, considerando apenas uma ocorrência por nome completo
function contarOcorrencias(planilha) {
    const contagem = {};

    planilha.forEach(item => {
        const nomeDriver = item.driver.trim(); // Nome completo com sufixo

        if (nomeDriver) {
            const nomeBase = nomeDriver.split(" ")[0]; // Pega apenas o nome base (ex: "Kennedy")

            // Adiciona o nome base ao objeto de contagem se ainda não estiver presente
            if (!contagem[nomeBase]) {
                contagem[nomeBase] = new Set(); // Usar um Set para evitar duplicatas
            }
            contagem[nomeBase].add(nomeDriver); // Adiciona o nome com sufixo ao Set
        }
    });

    // Transformar Set em contagem total
    for (const nomeBase in contagem) {
        contagem[nomeBase] = contagem[nomeBase].size; // Substitui pelo tamanho do Set
    }

    return contagem;
}

function exibirContagem(contagem) {
    const contagemDiv = document.getElementById("contagem-nomes");
    contagemDiv.innerHTML = ""; // Limpa resultados anteriores

    // Cria um contêiner flexível para os itens
    const contenedorFlex = document.createElement("div");
    contenedorFlex.classList.add("contagem-container");

    // Exibe todos os nomes e suas contagens
    for (const [nome, quantidade] of Object.entries(contagem)) {
        const p = document.createElement("p");
        p.classList.add("contagem-item"); // Adiciona uma classe para o estilo

        p.innerHTML = `<strong>${nome}</strong>: <span class="total">${quantidade}</span>`; 
        contenedorFlex.appendChild(p);
    }

    contagemDiv.appendChild(contenedorFlex); // Adiciona o contêiner flexível ao DOM
}

// Função para buscar pelo código
function buscarPorCodigo() {
    const codigoInput = document.getElementById("codigo").value.trim();
    const resultadosDiv = document.getElementById("resultado");
    resultadosDiv.innerHTML = ""; // Limpa resultados anteriores

    if (codigoInput === "") {
        resultadosDiv.innerHTML = "Por favor, insira um código para buscar.";
        return;
    }

    // Busca pelo código
    const resultado = dadosPlanilha.find(item => item.codigo.toString() === codigoInput);

    if (resultado) {
        const nomeDriver = resultado.driver.trim();
        resultadosDiv.innerHTML = `Código: <strong>${resultado.codigo}</strong> - Driver: <span class="highlight">${resultado.driver}</span>`;

        historicoBuscas.unshift({
            codigo: resultado.codigo,
            driver: resultado.driver
        });
        atualizarHistorico();

        speakText(`${resultado.driver}`);

        // Chama a função de impressão
        imprimirDriver(resultado.driver); // Chama a nova função de impressão

        const nomeBase = nomeDriver.split(" ")[0]; // Pega o nome base do driver

        // Verifica se o código já foi processado antes
        if (!codigosProcessados[codigoInput] && nomeContagemTotal[nomeBase] > 0) {
            nomeContagemTotal[nomeBase]--; // Subtrai 1 do total
            codigosProcessados[codigoInput] = true; // Marca o código como processado
            exibirContagem(nomeContagemTotal); // Atualiza a exibição da contagem
        }

    } else {
        resultadosDiv.innerHTML = "Nenhum resultado encontrado.";
    }

    // Limpa o campo de busca e mantém o foco
    document.getElementById("codigo").value = "";
    document.getElementById("codigo").focus();
}

function imprimirDriver(driver) {
    // Criar um iframe para impressão
    const iframe = document.createElement('iframe');
    iframe.style.position = 'absolute';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = 'none';
    document.body.appendChild(iframe);

    // Criar o conteúdo do iframe
    const estiloEtiquetadora = `
        <style>
            body {
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
                font-family: "Arial", sans-serif;
                overflow: hidden; /* Impede rolagem */
                background-color: white; /* Garante fundo branco */
            }
            h1 {
                font-size: 30px; /* Tamanho grande para a etiqueta */
                font-weight: bold;
                text-align: center;
                border: 2px solid black; /* Simula borda de etiqueta */
                padding: 20px;
                width: 60%; /* Controla o tamanho da etiqueta */
                margin: 0; /* Remove margem para centrar */
            }
            @media print {
                body {
                    width: 100%; /* Controla a largura da impressão */
                    height: 100%; /* Controla a altura da impressão */
                }
            }
        </style>
    `;

    const conteudo = `${estiloEtiquetadora}<h1>${driver}</h1>`;
    const doc = iframe.contentWindow.document;
    doc.open();
    doc.write(conteudo);
    doc.close();

    // Chama a impressão
    iframe.contentWindow.focus();
    iframe.contentWindow.print();

    // Remove o iframe após a impressão
    iframe.parentNode.removeChild(iframe);
}

function speakText(text) {
    const synth = window.speechSynthesis;
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = 'pt-BR'; 
    utterance.rate = 2.0;
    utterance.pitch = 1.3; 

    // Fala o texto
    synth.speak(utterance);
}

function atualizarHistorico() {
    const listaHistorico = document.getElementById("lista-historico");
    listaHistorico.innerHTML = "";

    historicoBuscas.forEach(item => {
        const li = document.createElement("li");
        li.innerHTML = `Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>`;
        listaHistorico.appendChild(li);
    });
}

// Função para exportar o histórico para um arquivo XLSX
function exportarHistorico() {
    const workbook = XLSX.utils.book_new();
    const worksheetData = [
        ["Código", "Driver"],
        ...historicoBuscas.map(item => [item.codigo, item.driver])
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Historico");

    // Gera o arquivo XLSX
    XLSX.writeFile(workbook, "historico.xlsx");
}

// Botão para reiniciar o histórico de buscas
document.getElementById("reiniciar-historico").addEventListener("click", function() {
    historicoBuscas = [];
    atualizarHistorico();
});

// Mostra um alerta de confirmação ao sair
window.onbeforeunload = function() {
    return "Tem certeza que deseja sair?";
};
