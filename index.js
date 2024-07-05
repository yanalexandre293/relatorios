const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');
const sizeOf = require('image-size');
const readline = require('readline-sync');
const csv = require('csv-parser');

function getUserInput(prompt) {
    return readline.question(prompt);
}

const municipio = getUserInput("MUNICIPIO: ").toUpperCase();
const serie = getUserInput("SERIE: ");
const codAvaliacao = getUserInput("CODIGO DE AVALIACAO: ");
const disciplina = getUserInput("DISCIPLINA (lp ou mt): ").toLowerCase();
const codCaderno = getUserInput("CADERNO: ");
const itens = [];
const itens1 = [];


// Função para carregar os dados do CSV
function carregarDadosCSV(callback) {
    let itensPlanilha = ""
    if (serie == '1' || serie == '2' || serie == '3') {
        itensPlanilha = "itens_1ao3_";
    }

    if (serie == '4' || serie == '5' || serie == '6') {
        itensPlanilha = "itens_4ao6_";
    }

    if (serie == '7' || serie == '8' || serie == '9') {
        itensPlanilha = "itens_7ao9_";
    }



    fs.createReadStream(`C:/Users/antoniomaluf/Documents/inDICA/itens/caderno${codCaderno}/itens${disciplina}.csv`, { encoding: 'latin1' })
        .pipe(csv({ separator: ';' }))
        .on('data', (row) => {
            const uppercaseRow = {};
            Object.keys(row).forEach(key => {
                uppercaseRow[key] = row[key].toUpperCase();
            });
            itens.push(uppercaseRow);
        });

    fs.createReadStream(`C:/Users/antoniomaluf/Documents/inDICA/output/AV_${codAvaliacao}/itens.csv`)
        .pipe(csv({ separator: ';' }))
        .on('data', (row) => {
            switch(row.dificuldade){
                case '1': row.dificuldade = 'Baixa'; break;
                case '2': row.dificuldade = 'Moderada'; break;
                case '3': row.dificuldade = 'Moderada'; break;
                case '4': row.dificuldade = 'Avançada'; break;
                default: row.dificuldade = 'Moderada'; break;
            }
            if(row.serie == serie && row.disciplina == '2' && disciplina == 'lp')  itens1.push(row);
            if(row.serie == serie && row.disciplina == '1' && disciplina == 'mt')  itens1.push(row);
        })
        .on('end', () => {
            console.log('Dados itens municipio: ', itens1);
            callback();
        });
}

// Função para preencher e gerar o documento
function gerarDocumento() {
    // Carregar o arquivo template
    const content = fs.readFileSync(`modelo${disciplina}.docx`, 'binary');

    // Criar um objeto zip a partir do conteúdo do template
    const zip = new PizZip(content);

    const imageOptions = {
        getImage(tagValue, tagName) {
            return fs.readFileSync(tagValue);
        },
        getSize(img, tagValue, tagName) {
            if (tagName.includes("imgItem")) {
                const dimensions = sizeOf(tagValue);
                console.log(dimensions);
                return [dimensions.width, dimensions.height];
            }
            if (tagName.includes("graficoPizza")) {
                return [500, 500];
            }
            return [500, 500];
        }
    };

    // Dados a serem preenchidos no template
    const data = {
        municipio: municipio,
        serie: serie,
    };


    itens.forEach((item, index) => {
        const itemNumber = index + 1;
        data[`bncc${itemNumber}`] = item.bncc;
        data[`bncctxt${itemNumber}`] = item.habilidade;
        data[`unidadeTematica${itemNumber}`] = item.genero;
        data[`objetoConhecimento${itemNumber}`] = item.objeto_conhecimento;
        data[`comentario${itemNumber}`] = item.comentario;
        data[`imgItem${itemNumber}`] = `q1.png`;
        data[`graficoPizza${itemNumber}`] = `C:/Users/antoniomaluf/Documents/inDICA/output/AV_${codAvaliacao}/AV_${codAvaliacao}_${serie}${disciplina}/grafitens/${serie}${disciplina}_item${itemNumber}.png`;
    });


    itens1.forEach((item, index) => {
        const itemNumber = index + 1;
        data[`dif${itemNumber}`] = item.dificuldade;
        data[`peso${itemNumber}`] = item.peso;
        data[`gab${itemNumber}`] = item.gabarito;
    });


    // Criar o Docxtemplater
    const doc = new Docxtemplater()
        .attachModule(new ImageModule(imageOptions))
        .loadZip(zip)
        .setData(data);

    // Renderizar o documento com os dados fornecidos
    doc.render();

    // Gerar o buffer do documento final
    const buf = doc.getZip().generate({
        type: 'nodebuffer',
        compression: 'DEFLATE',
    });

    // Salvar o documento gerado
    fs.writeFileSync(Buffer.from(`${municipio}_${serie}${disciplina}.docx`, 'latin1'), buf);

    console.log(`Relatório de ${disciplina}, ${serie} Ano de ${municipio}, gerado com sucesso!`);
}

// Chamada para carregar os dados do CSV e depois gerar o documento
carregarDadosCSV(gerarDocumento);
