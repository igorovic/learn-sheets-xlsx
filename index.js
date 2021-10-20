const xlsx = require('xlsx')

const filePath = "./data/perso.xlsx"
const wb = xlsx.readFile(filePath, { type: "file" })
const firstSheet = wb.SheetNames[0]

const ws = wb.Sheets[firstSheet]

// header: 0 indicates that first row actually contains headers
console.log(xlsx.utils.sheet_to_json(ws, { header: 0 }))

console.log("range", ws["!ref"])

console.log(xlsx.utils.decode_range("A1:F1"))
console.log(xlsx.utils.decode_range("A1:A2"))


const make_cols = refstr => {
    let o = [], C = xlsx.utils.decode_range(refstr).e.c + 1;
    for (var i = 0; i < C; ++i) o[i] = { name: xlsx.utils.encode_col(i), key: i }
    return o;
};


console.log(make_cols('A1:F2'))