/**  Node.js é um ambiente de tempo de execução JavaScript de código aberto e plataforma cruzada que
 *  também pode ser usado para ler de um arquivo e gravar em um arquivo. Esse arquivo  pode estar no formato txt, ods, xlsx, docx, etc.
 * */

/**O codigo a seguir retrata mais ou menos como um arquivo do Excel (.xlsx) é lido de um arquivo do Excel e, em seguida,
 *  convertido em JSON e também para gravar nele. Isso pode ser alcançado usando um pacote chamado **xlsx** para atingir nosso objetivo.
 *  */

import xlsx, { WorkBook } from "xlsx";
import path from 'path';

//pegando caminho do arquivo 
const caminho = path.resolve('./assets/estoque.xlsm')
const file = xlsx.readFile(caminho);

//criando vetores de armazenamento
let data: Array<WorkBook> = [];
const sheets: string[] = file.SheetNames


console.log(sheets)
for(let i = 0; i < sheets.length; i++){
   const temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
   temp.forEach((res:any) => {
      data.push(res)
   })
}
  
// Printing data
console.log(data);

//ref:: https://acervolima.com/como-ler-e-escrever-arquivos-do-excel-em-node-js/
// https://xlsxwriter.readthedocs.io/workbook.html