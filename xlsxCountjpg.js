const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Função para contar arquivos .jpg em um diretório
function countJpgFiles(directoryPath) {
    return new Promise((resolve, reject) => {
        fs.readdir(directoryPath, (err, files) => {
            if (err) {
                reject(`Erro ao ler o diretório: ${err}`);
                return;
            }

            let jpgCount = 0;

            files.forEach((file) => {
                if (path.extname(file).toLowerCase() === '.jpg') {
                    jpgCount++;
                }
            });

            resolve(jpgCount);
        });
    });
}

// Função para processar todos os subdiretórios
async function processDirectory(directoryPath) {
    const results = [];

    try {
        const items = await fs.promises.readdir(directoryPath, { withFileTypes: true });

        for (const item of items) {
            if (item.isDirectory()) {
                const subdirectoryPath = path.join(directoryPath, item.name);
                const jpgCount = await countJpgFiles(subdirectoryPath);
                results.push({ directory: item.name, jpgCount });
                console.log(`Diretório: ${item.name}, Quantidade de arquivos .jpg: ${jpgCount}`);
            }
        }
    } catch (err) {
        console.error(`Erro ao processar o diretório: ${err}`);
    }

    return results;
}

// Função para exportar os resultados para um arquivo Excel
function exportToExcel(data, outputPath) {
    const workbook = xlsx.utils.book_new();
    const worksheetData = [['Diretório', 'Quantidade de arquivos .jpg']];

    data.forEach((item) => {
        worksheetData.push([item.directory, item.jpgCount]);
    });

    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Resultados');
    xlsx.writeFile(workbook, outputPath);
}

// Diretório que você quer verificar
//const directoryPath = '../../../../../adda';
const directoryPath = 'P:\\imagens-e-commerce';
const outputPath = './listagemDeImagens.xlsx';

processDirectory(directoryPath)
    .then((results) => {
        exportToExcel(results, outputPath);
        console.log(`Os resultados foram exportados para ${outputPath}`);
    })
    .catch((error) => {
        console.error(error);
    });



    /** @macro VBA
     * 
     * Sub ColorirLinhasIguais()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim rng1 As Range
    Dim cell As Range
    Dim matchCell As Range

    ' Definir as planilhas
    Set ws1 = ThisWorkbook.Sheets("Planilha1") ' Altere "Planilha1" para o nome da sua primeira planilha
    Set ws2 = ThisWorkbook.Sheets("Planilha2") ' Altere "Planilha2" para o nome da sua segunda planilha

    ' Definir o intervalo da coluna A na primeira planilha
    Set rng1 = ws1.Range("A1:A" & ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row)

    ' Loop através de cada célula na coluna A da primeira planilha
    For Each cell In rng1
        ' Procurar o valor da célula na coluna A da segunda planilha
        Set matchCell = ws2.Range("A:A").Find(What:=cell.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Se encontrar o valor, mudar a cor da linha correspondente na segunda planilha
        If Not matchCell Is Nothing Then
            ws2.Rows(matchCell.Row).Interior.Color = RGB(255, 255, 0) ' Cor amarela
        End If
    Next cell
End Sub

     * 
     * 
     * 
     */