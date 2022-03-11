const getWS = (wb, indexOfSheet) => {
    let sheet_name = wb.SheetNames[indexOfSheet];
    return wb.Sheets[sheet_name];
}

// nom des colonnes
const columsNames = {
    mcode: 'M-CODE',
    repas: 'NOMBRE DE REPAS',
    number: 'Numbering Agent',
    shift: 'Shift Name',
    transpday: 'TRANSPORT JOUR',
    transpnight: 'TRANSPORT SOIR',
    transp: 'TRANSPORT'
}

// read without style
const readWBxlsx = (filename) => {
    const xlsx = require('xlsx');
    return xlsx.readFile(filename, {cellDates: true});
}

// read with style
const readWBxlsxstyle = (filename) => {
    const xlsxstyle = require('xlsx-style');
    return xlsxstyle.readFile(filename, {cellStyles: true});
}

const findData = (ws, key) => {
    let data = [];
    Object.keys(ws).forEach(e => {
        if (ws[e].v && new String(ws[e].v).match(key))
            data.push({c: e, v: ws[e]});
    })
    return data;
}

const combineStyle = (wb_xlsx, wb_xlsx_style) => {
    let sheets_leng = wb_xlsx.SheetNames.length;
    for (let i = 0; i < sheets_leng; i++) {
        let ws = wb_xlsx.Sheets[wb_xlsx.SheetNames[i]];
        let ws_s = wb_xlsx_style.Sheets[wb_xlsx_style.SheetNames[i]];
        Object.keys(ws).forEach(key => {
            if (typeof ws_s[key] !== 'undefined') {
                let s = ws_s[key].s;
                ws_s[key] = ws[key];
                ws_s[key].s = s;
            }
        });
    }
    return wb_xlsx_style;
}


// arrange transport
const arrangeTRANSPORTS = (ws) => {
    let data = [];
    let data_trans = findData(ws, columsNames.transp).map(e => e.c);
    let temp = null;
    data_trans.forEach(e => {
        data.push(e);
        if (temp !== null) {
            if (new String(e).substring(0, 1) === 'N' && new String(e).substring(1, 2) === new String(temp).substring(1, 2)) {
                let data_comb = [temp, e];
                data[data.indexOf(temp)] = data_comb;
                data.pop();
            }
        }
        temp = e;
    });
    return data;
}



const getColumnName = (ws, columnName) => {
    return findData(ws, columnName).map(e => e.c);
}

const getGroupedRequiredCol = (ws) => {
    // group column
    let MCODE_col = getColumnName(ws, 'M-CODE');
    let NUMBERINGAGENT_col = getColumnName(ws,'Numbering Agent');
    let NOMBREREPAS_col = getColumnName(ws,'NOMBRE DE REPAS');
    let TRANSPORT_col = arrangeTRANSPORTS(ws);
    return groupCol(MCODE_col, NUMBERINGAGENT_col, NOMBREREPAS_col, TRANSPORT_col);
}

const groupCol = (...cols) => {
    let data = [];
    for (let i = 0; i < cols[0].length; i++) {
        let coldata = [];
        cols.forEach(el => {
            if (el[i].constructor === Array) {
                el[i].forEach(e => {
                    coldata.push(e);
                })
            } else {
                coldata.push(el[i]);
            }
        })
        data.push(coldata)
    }
    return data;
}

// afficher toutes les données de chaque colonne
const fetchData = (ws, group_col_data = []) => {
    const xlsx = require('xlsx');
    let jsonArray = [];
    if (group_col_data.length === 0)
        group_col_data = getGroupedRequiredCol(ws);

    group_col_data.forEach(array => {
        let line = parseInt(array[1].substring(1, array[1].length)) + 1;
        let letter0 = array[0].substring(0, 1); // M-CODE
        let letter1 = array[1].substring(0, 1); // Numbering Agent
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (typeof ws[letter1+line] !== 'undefined' && typeof ws[letter1+line] !== 'undefined') {
            // objet pour construire un élément.
            let obj = {};
            // creer un key shift sur l'objet.
            let shift = ws[letter0+(parseInt(array[0].substring(1, array[0].length)) - 1)];
            if (shift) obj[columsNames.shift] = shift.v;
            // parcourir le tableau cols
            array.forEach(val => {
                let letter_i = val.substring(0, 1);
                // si la valeur de la colonne Numbering Agent n'est pas vide.
                if (ws[letter_i+line]) {
                    obj[ws[val].w] = ws[letter_i+line].w || '';
                } else {
                    obj[ws[val].w] = ''; // definir vide par défaut si la cellule est vide.
                }
            })
            // ajouter dans le tableau l'objet qu'on a crée.
            jsonArray.push(obj);
            // passer dans la ligne suivante.
            line++;
        }
    });
    // retourner la le tableau.
    return jsonArray;
}

const createOutput = (DATA_RH = [], wb, wb_style) => {
    var agent_found = 0;
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    // parcourir tous les feuilles SHEETS
    for (let i = 0; i < newWorkbook.SheetNames.length; i++) {
        let ws = newWorkbook.Sheets[newWorkbook.SheetNames[i]];
        // chercher ou se situe le 2000 et 1000
        let colsToFill = Array.from(Object.keys(ws)).filter(v => ws[v].w === '1000' || ws[v].w === '2000' || ws[v].w === 'TRANSPORT (2000/day)' || new String(ws[v].w).match('REPAS'));
            colsToFill = colsToFill.map(e => { return {c: e, v: ws[e]} });
        // total transport
        // let total_col = Array.from(Object.keys(ws)).find(v => new String(ws[v].w).match('TOTAL'));
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numberingagent = ws[important_cols[0]+line].w;
                let mcode = ws[important_cols[1]+line].w;
                let info = null;
                // PDFButler
                if (new String(numberingagent).match('PDFB-')) {
                    info = DATA_RH.find(e => e[columsNames.number] === numberingagent);
                } else {
                    // get info via RH by M-CODE and Numbering Agent
                    info = DATA_RH.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                }
                if (info) {
                    // increment the numbers of agent
                    agent_found += 1;
                    switch (info[columsNames.shift]) {
                        case 'SHIFT 3' : case 'SHIFT WEEKEND':
                            // 1000
                            let col1000 = colsToFill.find(e => e.v.v === 1000);
                            if (col1000) {
                                if (columsNames.transpnight in info) {
                                    ws[col1000.c.substring(0, 1)+line].v = parseFloat(info[columsNames.transpnight]) || 0;
                                    ws[col1000.c.substring(0, 1)+line].w = info[columsNames.transpnight];
                                }
                            }
                            // 2000
                            let col2000 = colsToFill.find(e => e.v.v === 2000);
                            if (col2000) {
                                if (columsNames.transpday in info) {
                                    ws[col2000.c.substring(0, 1)+line].v = parseFloat(info[columsNames.transpday]) || 0;
                                    ws[col2000.c.substring(0, 1)+line].w = info[columsNames.transpday];
                                }
                            }
                            
                        break;
                        default: 
                            
                            if (ws[colsToFill[0].c.substring(0, 1)+line]) {
                                ws[colsToFill[0].c.substring(0, 1)+line].v = parseFloat(info[columsNames.transp]) || 0;
                                ws[colsToFill[0].c.substring(0, 1)+line].v = info[columsNames.transp];
                            }
                            break;
                    }
                    
                    // Nombre de repas
                    colRep = colsToFill.find(e => new String(e.v.v).toUpperCase().match('REPAS'));
                    if (colRep)
                        if (columsNames.repas in info) {
                            if (typeof ws[colRep.c.substring(0, 1)+line] === 'undefined')
                                ws[colRep.c.substring(0, 1)+line] = {};
                            ws[colRep.c.substring(0, 1)+line].t = 'n';
                            let number = (parseFloat(info[columsNames.repas]) || 0) * 3500;
                            ws[colRep.c.substring(0, 1)+line].v = number;
                            ws[colRep.c.substring(0, 1)+line].w = ws[colRep.c.substring(0, 1)+line].v;
                    }
                }
            }
            line ++;
        }
    }

    return {
        agent_found : agent_found,
        wb: combineStyle(newWorkbook, wb_style)
    };
    
}

// sauvegarder le fichier xlsx
const saveFile = (wb, filename) => {
    let save = require('sheetjs-style');
    save.writeFile(wb, filename, {type: 'file'});
}

function randomCode() {
    var code = "";
    let v = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ!é&#";
    for (let i = 0; i < 6; i++) { // 6 characters
      let char = v.charAt(Math.random() * v.length - 1);
      code += char;
    }
    return code;
}
// export function
module.exports = {readWBxlsx, readWBxlsxstyle, arrangeTRANSPORTS, getColumnName, groupCol, fetchData, createOutput, saveFile, getWS, getGroupedRequiredCol, randomCode};