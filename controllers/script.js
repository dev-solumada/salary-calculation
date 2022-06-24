const getWS = (wb, indexOfSheet) => {
    let sheet_name = wb.SheetNames[indexOfSheet];
    return wb.Sheets[sheet_name];
}

// get sheetname
const getSheetIndex = (wb, sheetname) => {
    let index = wb.SheetNames.indexOf(sheetname);
    return index > -1 ? index : null;
}

// nom des colonnes
const columsNames = {
    mcode: 'M-CODE',
    repas: 'NOMBRE DE REPAS',
    number: 'Numbering Agent',
    shift: 'Shift Name',
    transpday: 'TRANSPORT JOUR',
    transpnight: 'TRANSPORT SOIR',
    transp: 'TRANSPORT',
    salaryUP: 'SALARY UP',
    salaryAGROBOX: 'SALARY AGROBOX',
    salaryARCO: 'SALARY ARCO',
    salaryUPC: 'SALARY UPC',
    salaryJEFACTURE: 'SALARY JeFACTURE',
    salaryPWC: 'SALARY PWC',
    salarySPOTCHECK: 'SALARY SPOTCHECK',
    willemen: { 
        salary: 'SALARY WILLEMEN',
        ai: 'AI DOCUMENT',
        id: 'ID DOCUMENT',
        limosa: 'LIMOSA DOCUMENT',
    },
    vandorp: 'SALARY VANDORP'
};


const colsIndexNames = () => {
    let alphabet = String.fromCharCode(...Array(123).keys()).slice(97).toUpperCase();
    return alphabet;
}

// read without style
const readWBxlsx = (filename) => {
    const xlsx = require('xlsx');
    return xlsx.readFile(filename, {cellDates: true});
}

// read with style
const readWBxlsxstyle = (filename) => {
    const xlsx = require('xlsx-style');
    return xlsx.readFile(filename, {cellStyles: true});
}

const findData = (ws, key) => {
    let data = [];
    Object.keys(ws).forEach(e => {
        if (ws[e].v && new String(ws[e].v).match(key))
            data[data.length] = {c: e, v: ws[e]};
    })
    return data;
}

const combineStyle = (wb_xlsx, wb_xlsx_style) => {
    const XLSX = require('xlsx')
    let sheets_leng = wb_xlsx.SheetNames.length;
    for (let i = 0; i < sheets_leng; i++) {
        let ws = wb_xlsx.Sheets[wb_xlsx.SheetNames[i]];
        let ws_s = wb_xlsx_style.Sheets[wb_xlsx_style.SheetNames[i]];
        var range = XLSX.utils.decode_range(ws['!ref']);
        Object.keys(ws).forEach(key => {
            if (ws_s[key]) {
                let s = ws_s[key].s;
                ws_s[key] = ws[key];
                ws_s[key].s = s;
            }
        });

        /* ELIMINER LES BACKGROUND NOIR */
        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            // loo all cells in the current column
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
                // cell styled
                const cellStyled = ws_s[cellName];
                if (cellStyled && cellStyled.s) {
                    if (cellStyled.s.fill && cellStyled.s.fill.bgColor)
                        if (cellStyled.s.fill.bgColor.rgb === '000000') { // if bg is dark
                            cellStyled.s.fill.bgColor = {}; // set bg to white
                            delete cellStyled.s.fill;
                            ws_s[cellName].s = cellStyled.s;
                    }
                }
            }
        }
        
    }

    return wb_xlsx_style;
}


// arrange transport
const arrangeTRANSPORTS = (ws) => {
    let data = [];
    let data_trans = findData(ws, columsNames.transp).map(e => e.c);
    let temp = null;
    
    data_trans.forEach(e => {
        data[data.length] = e;
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
    let NUMBERINGAGENT_col = getColumnName(ws, columsNames.number);
    let NOMBREREPAS_col = getColumnName(ws, columsNames.repas);
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
                    coldata[coldata.length] = e;
                })
            } else {
                coldata[coldata.length] = el[i];
            }
        })
        data[data.length] = coldata;
    }
    return data;
}

// afficher toutes les données de chaque colonne
const fetchData = (ws, group_col_data = []) => {
    // const xlsx = require('xlsx');
    let jsonArray = [];
    if (group_col_data.length === 0)
        group_col_data = getGroupedRequiredCol(ws);

    group_col_data.forEach(array => {
        let line = parseInt(array[1].substring(1, array[1].length)) + 1;
        let letter0 = array[0].substring(0, 1); // M-CODE
        let letter1 = array[1].substring(0, 1); // Numbering Agent
        // const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (ws[letter0+line] || ws[letter1+line]) {
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
            jsonArray[jsonArray.length] = obj;
            // passer dans la ligne suivante.
            line++;
        }
    });
    // retourner la le tableau.
    return jsonArray;
}


const outTR = (DATA_RH = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    // parcourir tous les feuilles SHEETS
    for (let i = 0; i < newWorkbook.SheetNames.length; i++) {
        let ws = newWorkbook.Sheets[newWorkbook.SheetNames[i]];
        // chercher ou se situe le 2000 et 1000
        let colsToFill = Array.from(Object.keys(ws)).filter(v => ws[v].w === getVar().transpprice.night
            || ws[v].w === getVar().transpprice.day
            || ws[v].w === 'TRANSPORT (2000/day)'
            || new String(ws[v].w).match('REPAS'));
        colsToFill = colsToFill.map(e => { return {c: e, v: ws[e]} });
        // total transport
        // let total_col = Array.from(Object.keys(ws)).find(v => new String(ws[v].w).match('TOTAL'));
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        let target100 = null, 
            target200 = null;
        while (line <= rows.length) {
            if (ws[important_cols[0] + line] !== undefined || ws[important_cols[1] + line] !== undefined) {
                // numbering agent
                let numberingagent = new String(ws[important_cols[0]+line]?.w ?? '').trim();
                let mcode = new String(ws[important_cols[1]+line]?.w ?? '').trim();
                let info = null;

                // PDFButler
                if (new String(numberingagent).match('PDFB-')) {
                    info = DATA_RH.find(e => e[columsNames.number] === numberingagent);
                } else if (new String(numberingagent).includes('GARDIEN CHARLES')) {
                    info = DATA_RH.find(e => e[columsNames.mcode] === 'Gardien');
                } else {
                    // get info via RH by M-CODE and Numbering Agent
                    info = DATA_RH.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                }
                if (info) {
                    // increment the numbers of agent
                    var col2000 = colsToFill.find(e => e.v.v === parseFloat(getVar().transpprice.day));
                    var col1000 = colsToFill.find(e => e.v.v === parseFloat(getVar().transpprice.night));
                    if (col1000) target100 = col1000.c.substring(0, 1)+line;
                    if (col2000) target200 = col2000.c.substring(0, 1)+line;
                    switch (info[columsNames.shift]) {
                        case 'SHIFT 3' : case 'SHIFT WEEKEND':
                            // 1000
                            if (col1000) {
                                if (columsNames.transpnight in info) {
                                    ws[target100].v = parseFloat(info[columsNames.transpnight]) || 0;
                                    ws[target100].w = info[columsNames.transpnight];
                                }
                            }
                            // 2000
                            if (col2000) {
                                if (columsNames.transpday in info) {
                                    ws[target200].v = parseFloat(info[columsNames.transpday]) || 0;
                                    ws[target200].w = info[columsNames.transpday];
                                }
                            }
                           
                            
                        break;
                        default: 
                            // for gardien
                            if (info[columsNames.mcode] === 'Gardien') {
                                ws['I'+line].v = parseFloat(info[columsNames.transp]) || 0;
                                ws['I'+line].w = info[columsNames.transp];
                            }
                            else if (ws[colsToFill[0].c.substring(0, 1)+line]) {
                                ws[colsToFill[0].c.substring(0, 1)+line].v = parseFloat(info[columsNames.transp]) || 0;
                                ws[colsToFill[0].c.substring(0, 1)+line].w = info[columsNames.transp];
                            }
                
                            break;
                    } 

                    if (col1000 && col2000) {
                        // set formule
                        let ciNames = colsIndexNames();
                        let indexOf100 = ciNames.indexOf(col1000.c.substring(0, 1));
                        let nextCol = ciNames[indexOf100 + 1];
                        if (nextCol.length === 1) {
                            if (target100 !== null  && target200 !== null) {
                                if (info[columsNames.mcode] !== 'Gardien')
                                    ws[nextCol+line].f = `${target200}*${col2000.c}+${target100}*${col1000.c}`;
                            }
                        } 
                    }
                    
                    // Nombre de repas
                    colRep = info[columsNames.mcode] !== 'Gardien' ? 
                        colsToFill.find(e => new String(e.v.v).toUpperCase().match('REPAS'))
                        : {c: 'G43'};
                    
                    if (colRep) {
                        if (columsNames.repas in info) {
                            if (typeof ws[colRep.c.substring(0, 1)+line] === 'undefined')
                                ws[colRep.c.substring(0, 1)+line] = {};
                            ws[colRep.c.substring(0, 1)+line].t = 'n';
                            let number = info[columsNames.mcode] === 'Gardien' ? (parseFloat(info[columsNames.repas]) || 0) : (parseFloat(info[columsNames.repas]) || 0) * parseFloat(getVar().repas);
                            ws[colRep.c.substring(0, 1)+line].v = number;
                            ws[colRep.c.substring(0, 1)+line].w = ws[colRep.c.substring(0, 1)+line].v;
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
    
}

// sauvegarder le fichier xlsx
const saveFile = (wb, wbs, filename) => {
    let save = require('sheetjs-style');
    save.writeFile(combineStyle(wb, wbs), filename, {type: 'file'});
}

function randomCode(length = 6) {
    var code = "";
    let v = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ!é&#";
    for (let i = 0; i < length; i++) { // 6 characters
      let char = v.charAt(Math.random() * v.length - 1);
      code += char;
    }
    return code;
}

// RANDOM NUMBER CODE
function randomnNumberCode(length = 6) {
    var code = "";
    let v = "0123456789";
    for (let i = 0; i < length; i++) { // 6 characters
      let char = v.charAt(Math.random() * v.length - 1);
      code += char;
    }
    return code;
}

// ===========================================================
/**
 * SALARY EXTRACTION
 */
// ===========================================================

const outUp = (DATA_RH = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // parcourir tous les feuilles SHEETS
    for (let i = 0; i < newWorkbook.SheetNames.length; i++) {
        let ws = newWorkbook.Sheets[newWorkbook.SheetNames[i]];
        // target column to salary up
        let colGSS = sheetColumn.gss['sheet'+(i+1)];
        let colSalaryUp = colGSS ? (colGSS['up'] || null) : null;
        
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numberingagent = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via RH by M-CODE and Numbering Agent
                let info = DATA_RH.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                if (info) {
                    // salary 
                    if (columsNames.salaryUP in info) {
                        // cols to fill
                        if (colSalaryUp) {
                            let colIndex = colSalaryUp+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.salaryUP];
                            ws[colIndex].w = new String(info[columsNames.salaryUP]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
}

const fetchUPData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        for (let colNum = range.s.c; colNum <= 3; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('A')) obj[columsNames.number] = cell.w;
                if (cellName.includes('B')) obj[columsNames.mcode] = cell.w;
                if (cellName.includes('C')) obj[columsNames.salaryUP] = parseFloat(cell.w) || 0;
            }
        // NOTE: secondCell is undefined if it does not exist (i.e. if its empty)
        }
        // if keys exist
        if (columsNames.mcode in obj && columsNames.salaryUP in obj) 
            data[data.length] = obj;
    }
    return data;
}

// ===========================================================
/**
 * AGROBOX
 */
// ===========================================================

const outAGROBOX = (DATA_RH = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // parcourir tous les feuilles SHEETS
    for (let i = 0; i < newWorkbook.SheetNames.length; i++) {
        let ws = newWorkbook.Sheets[newWorkbook.SheetNames[i]];
        // target column to salary agrobox
        let colGSS = sheetColumn.gss['sheet'+(i+1)];
        let colSalaryAgrobox = colGSS ? (colGSS['agrobox'] || null) : null;
        
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numberingagent = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via RH by M-CODE and Numbering Agent
                let info = DATA_RH.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                if (info) {
                    // salary 
                    if (columsNames.salaryAGROBOX in info) {
                        // cols to fill
                        if (colSalaryAgrobox) {
                            let colIndex = colSalaryAgrobox+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.salaryAGROBOX];
                            ws[colIndex].w = new String(info[columsNames.salaryAGROBOX]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
    
}

const fetchAgroboxData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        for (let colNum = range.s.c; colNum <= 3; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('A')) obj[columsNames.number] = cell.w;
                if (cellName.includes('B')) obj[columsNames.mcode] = cell.w;
                if (cellName.includes('C')) obj[columsNames.salaryAGROBOX] = parseFloat(cell.v) || 0;
            }
        // NOTE: secondCell is undefined if it does not exist (i.e. if its empty)
        }
        // if keys exist
        if ((columsNames.mcode in obj || columsNames.number in obj) && columsNames.salaryAGROBOX in obj) 
            data[data.length] = obj;
    }
    return data;
}

// ===========================================================
/**
 * ACRO
 */
// ===========================================================

const outARCO = (ARCO_Data = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // loop through agents to find sheets name by numbering
    for (let i = 0; i < ARCO_Data.length; i++) {
        let dataItem = ARCO_Data[i];
        // find sheet by agentnumbering
        let sheetName = newWorkbook.SheetNames.find(e => e?.substring(0, 2) === dataItem[columsNames.number]?.substring(0, 2));
        if (!sheetName) continue;
        let ws = newWorkbook.Sheets[sheetName];
        if (!ws) continue;
        // target column to salary pwc
        let colGSS = sheetColumn.gss['sheet'+(newWorkbook.SheetNames.indexOf(sheetName) + 1)];
        let colSalaryArco = colGSS ? (colGSS['arco'] || null)  : null;
        
        // total transport
        let important_cols = ['A', 'B'];
        let line = 0;
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numberingagent = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via RH by M-CODE and Numbering Agent
                let info = ARCO_Data.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                if (info) {
                    // salary 
                    if (columsNames.salaryARCO in info) {
                        // cols to fill
                        if (colSalaryArco) {
                            let colIndex = colSalaryArco+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.salaryARCO];
                            ws[colIndex].w = new String(info[columsNames.salaryARCO]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
    
}

const fetchArcoData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('J')) obj[columsNames.number] = cell.w;
                if (cellName.includes('I')) obj[columsNames.mcode] = cell.w;
                if (cellName.includes('L')) {
                    obj[columsNames.salaryARCO] = Math.round(parseFloat(cell.v)) || 0;
                }
            }
        }
        // if keys exist
        if ((columsNames.mcode in obj || columsNames.number in obj) && columsNames.salaryARCO in obj) 
            data[data.length] = obj;
    }
    return data;
}


// ===========================================================
/**
 * UPC
 */
// ===========================================================
 const fetchUPCData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        for (let colNum = range.s.c; colNum <= 3; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('A')) obj[columsNames.number] = cell.w;
                if (cellName.includes('B')) obj[columsNames.mcode] = cell.w;
                if (cellName.includes('C')) obj[columsNames.salaryUPC] = parseFloat(cell.w) || 0;
            }
        }
        // if keys exist
        if (columsNames.mcode in obj && columsNames.salaryUPC in obj) 
            data[data.length] = obj;
    }
    return data;
}

const outUPC = (UPC_Data = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // loop through agents to find sheets name by numbering
    for (let i = 0; i < UPC_Data.length; i++) {
        let dataItem = UPC_Data[i];
        // find sheet by agentnumbering
        let sheetName = newWorkbook.SheetNames.find(e => e?.substring(0, 2) === dataItem[columsNames.number]?.substring(0, 2));
        if (!sheetName) continue;
        let ws = newWorkbook.Sheets[sheetName];
        if (!ws) continue;
        // target column to salary pwc
        let colGSS = sheetColumn.gss['sheet'+(newWorkbook.SheetNames.indexOf(sheetName) + 1)];
        let colSalaryUPC = colGSS ? (colGSS['upc'] || null) : null;
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numberingagent = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via RH by M-CODE and Numbering Agent
                let info = UPC_Data.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                if (info) {
                    // salary 
                    if (columsNames.salaryUPC in info) {
                        // cols to fill
                        if (colSalaryUPC) {
                            let colIndex = colSalaryUPC+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.salaryUPC];
                            ws[colIndex].w = new String(info[columsNames.salaryUPC]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
}


// convert date to [dd,mm,yyyy]
function getDateNow() {
    var date = new Date(Date.now()),
    mnth = ("0" + (date.getMonth() + 1)).slice(-2),
    day = ("0" + date.getDate()).slice(-2);
    return [day,  mnth, date.getFullYear()];
}

// delete file
const deleteFile = (filePath, ms) => {
    const fs = require('fs');
    setTimeout(() => {
        // check file if exists
        if (fs.existsSync(filePath))
            fs.unlinkSync(filePath);
    }, ms);
}


/**
 * Add some points to have a clear date
 * ####### => ##.##.####
 */
const getDateInFileName = (filename) => {
    try {
        let date = filename.split(' ')[0];
        return date.substring(0,2) +'.'+date.substring(2,4)+'.'+date.substring(4,date.length);   
    } catch {
        return ".";
    }
}

/**
GET file name
*/
const getFirstDateInOutputFilename = (filename = '') => {
    try {
        let splitted = filename.split('-');
        return splitted.length === 1 ? null : filename.split('-')[0];
    } catch {
        return null;
    }
}

const getVar = () => {
    const mongoose = require('mongoose');
    var json = require('./var.json');
    mongoose.connect(
        process.env.MONGO_URI,
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
    ).then(async () => {
        // get vars on the database
        let vars = await JSONVarSchema.findOne();
        json = JSON.parse(vars.json);
    }).catch(err => {
        json = require('./var.json');
    });
    return json;
}

const accessDB = async (func) => {
    require('mongoose').connect(
        process.env.MONGO_URI, {
        useUnifiedTopology: true,
        UseNewUrlParser: true
    }).then(async () => {
        return func();
    }).catch(err => {
        console.log(err);
    });
}

// ===========================================================
/**
 * JeFACTURE
 */
// ===========================================================

const fetchJEFACTUREData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        for (let colNum = range.s.c; colNum <= 3; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('A')) obj[columsNames.number] = cell.w;
                if (cellName.includes('B')) obj[columsNames.mcode] = cell.w;
                if (cellName.includes('C')) obj[columsNames.salaryJEFACTURE] = parseFloat(cell.w) || 0;
            }
        }
        // if keys exist
        if (columsNames.mcode in obj && columsNames.salaryJEFACTURE in obj) 
            data[data.length] = obj;
    }
    return data;
}

const outJEFACTURE = (JFACTURE_Data = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // loop through agents to find sheets name by numbering
    for (let i = 0; i < JFACTURE_Data.length; i++) {
        let dataItem = JFACTURE_Data[i];
        // find sheet by agentnumbering
        let sheetName = newWorkbook.SheetNames.find(e => e?.substring(0, 2) === dataItem[columsNames.number]?.substring(0, 2));
        if (!sheetName) continue;
        let ws = newWorkbook.Sheets[sheetName];
        if (!ws) continue;
        // target column to salary pwc
        let colGSS = sheetColumn.gss['sheet'+(newWorkbook.SheetNames.indexOf(sheetName) + 1)];
        let colSalaryJEFACTURE = colGSS ? (colGSS['jefacture'] || null) : null;
        
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numberingagent = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via RH by M-CODE and Numbering Agent
                let info = JFACTURE_Data.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                if (info) {
                    // salary 
                    if (columsNames.salaryJEFACTURE in info) {
                        // cols to fill
                        if (colSalaryJEFACTURE) {
                            let colIndex = colSalaryJEFACTURE+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.salaryJEFACTURE];
                            ws[colIndex].w = new String(info[columsNames.salaryJEFACTURE]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
}


// ===========================================================
/**
 * JeFACTURE
 */
// ===========================================================

const fetchPWCData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = 1; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        // column alpha name O, N
        for (let colNum = range.s.c; colNum <= 14; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('B')) obj[columsNames.number] = cell?.w || '';
                if (cellName.includes('C')) obj[columsNames.mcode] = cell?.w || '';
                if (cellName.includes('O')) obj[columsNames.salaryPWC] = parseFloat(cell?.w) || 0;
            }
        }
        // if keys exist
        if (columsNames.mcode in obj && columsNames.salaryPWC in obj)
            data[data.length] = obj;
    }
    return data;
}

const outPWC = (PWC_Data = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // loop through agents to find sheets name by numbering
    for (let i = 0; i < PWC_Data.length; i++) {
        let dataItem = PWC_Data[i];
        // find sheet by agentnumbering
        let sheetName = newWorkbook.SheetNames.find(e => e?.substring(0, 2) === dataItem[columsNames.number]?.substring(0, 2));
        if (!sheetName) continue;
        let ws = newWorkbook.Sheets[sheetName];
        if (!ws) continue;
        // target column to salary pwc
        let colGSS = sheetColumn.gss['sheet'+(newWorkbook.SheetNames.indexOf(sheetName) + 1)];
        let colSalaryPWC = colGSS ? (colGSS['pwc'] || null) : null;
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[1]+line]) {
                let number = new String(ws[important_cols[0]+line]?.w ?? '').trim();
                let mcode = new String(ws[important_cols[1]+line]?.w ?? '').trim();
                // get info via the PWC data by only M-CODE
                let info = PWC_Data.find(e => e[columsNames.mcode] === mcode && e[columsNames.number] === number);
                if (info) {
                    // salary 
                    if (columsNames.salaryPWC in info) {
                        // cols to fill
                        if (colSalaryPWC) {
                            let colIndex = colSalaryPWC+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.salaryPWC];
                            ws[colIndex].w = new String(info[columsNames.salaryPWC]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
}


// ===========================================================
/**
 * JeFACTURE
 */
// ===========================================================

const fetchSPOTCHECKData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = 2; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};// column alpha name O, N
        for (let colNum = 1; colNum <= range.e.c; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('B')) obj[columsNames.mcode] = cell?.w;
                if (cellName.includes('C')) obj[columsNames.number] = cell?.w;
                if (cellName.includes('J')) obj[columsNames.salarySPOTCHECK] =( parseFloat(cell?.w) || 0) * parseFloat(getVar().spotckeck_mult);
            }
        }
        // if keys exist
        if (columsNames.mcode in obj && columsNames.salarySPOTCHECK in obj) 
            data[data.length] = obj;
    }
    return data;
}

const outSPOTCHECK = (SPOTCHECK_Data = [], wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // loop through agents to find sheets name by numbering
    for (let i = 0; i < SPOTCHECK_Data.length; i++) {
        let dataItem = SPOTCHECK_Data[i];
        // find sheet by agentnumbering
        let sheetName = newWorkbook.SheetNames.find(e => e?.substring(0, 2) === dataItem[columsNames.number]?.substring(0, 2));
        if (!sheetName) continue;
        let ws = newWorkbook.Sheets[sheetName];
        if (!ws) continue;
        // target column to salary agrobox
        let colGSS = sheetColumn.gss['sheet'+(newWorkbook.SheetNames.indexOf(sheetName) + 1)];
        let colSalarySPOTCHECK = colGSS ? (colGSS['spotcheck'] || null) : null;
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numbering = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via the SPOTCHECK data by only M-CODE
                let info = SPOTCHECK_Data.find(e => e[columsNames.mcode] === mcode && e[columsNames.number] === numbering);
                if (info) {
                    // salary 
                    if (columsNames.salarySPOTCHECK in info) {
                        // cols to fill
                        if (colSalarySPOTCHECK) {
                            let colIndex = colSalarySPOTCHECK+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.salarySPOTCHECK];
                            ws[colIndex].w = new String(info[columsNames.salarySPOTCHECK]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
}

/**
 * FETCH WILLEMEN DATA
 * @param {Workbook} ws 
 * @returns Array
 */

const fetchWillemenData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var jsonVar = getVar();
    var range = XLSX.utils.decode_range(ws['!ref']);
    var [ai, id, limosa] = [
        new String(ws['D25']?.w).replace('MGA', ''),
        new String(ws['D26']?.w).replace('MGA', ''),
        new String(ws['D27']?.w).replace('MGA', '')
    ];
    for (let rowNum = 1; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        // column alpha name O, N
        for (let colNum = 0; colNum <= range.e.c; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('B')) obj[columsNames.number] = cell?.w;
                if (cellName.includes('C')) obj[columsNames.mcode] = cell?.w;
                if (cellName.includes('E')) obj[columsNames.willemen.ai] = cell?.v * (parseFloat(ai) || parseFloat(jsonVar.willemen.ai));
                if (cellName.includes('F')) obj[columsNames.willemen.id] = cell?.v * (parseFloat(id) || parseFloat(jsonVar.willemen.id));
                if (cellName.includes('G')) obj[columsNames.willemen.limosa] = cell?.v * (parseFloat(limosa) || parseFloat(jsonVar.willemen.limosa));
            }
        }
        // if keys exist
        if (columsNames.mcode in obj) {
            obj[columsNames.willemen.salary] = (obj[columsNames.willemen.ai] || 0) + (obj[columsNames.willemen.id] || 0) + (obj[columsNames.willemen.limosa] || 0);
            data[data.length] = obj;
        }
    }
    return data;
}  

const outWillemen = (WILLEMEN_Data = [],  wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // loop through agents to find sheets name by numbering
    for (let i = 0; i < WILLEMEN_Data.length; i++) {
        let dataItem = WILLEMEN_Data[i];
        // find sheet by agentnumbering
        let sheetName = newWorkbook.SheetNames.find(e => e?.substring(0, 2) === dataItem[columsNames.number]?.substring(0, 2));
        if (!sheetName) continue;
        let ws = newWorkbook.Sheets[sheetName];
        if (!ws) continue;
        // target column to salary agrobox
        let colGSS = sheetColumn.gss['sheet'+(newWorkbook.SheetNames.indexOf(sheetName) + 1)];
        let colSalaryWILLEMEN = colGSS ? (colGSS['willemen'] || null) : null;
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numbering = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via the WILLEMEN data by only M-CODE
                let info = WILLEMEN_Data.find(e => e[columsNames.mcode] === mcode && e[columsNames.number] === numbering);
                if (info) {
                    // salary 
                    if (columsNames.willemen.salary in info) {
                        // cols to fill
                        if (colSalaryWILLEMEN) {
                            let colIndex = colSalaryWILLEMEN+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.willemen.salary];
                            ws[colIndex].w = new String(info[columsNames.willemen.salary]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
}


/**
 * FETCH WILLEMEN DATA
 * @param {Workbook} ws 
 * @returns Array
 */

const fetchVanDorpData = (ws) => {
    const XLSX = require('xlsx');
    var data = [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (let rowNum = 1; rowNum <= range.e.r; rowNum++) {
        // loo all cells in the current column
        let obj = {};
        // column alpha name O, N
        for (let colNum = 0; colNum <= range.e.c; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            // cell styled
            const cell = ws[cellName];
            if (cell) {
                if (cellName.includes('B')) obj[columsNames.number] = cell?.w;
                if (cellName.includes('C')) obj[columsNames.mcode] = cell?.w;
                if (cellName.includes('P')) obj[columsNames.vandorp] = parseFloat(cell?.v) || 0;
            }
        }
        // if keys exist
        if (columsNames.mcode in obj) {
            data[data.length] = obj;
        }
    }
    return data;
}  

const outVanDorp = (VANDORP_Data = [],  wb) => {
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    const sheetColumn = getVar();
    // loop through agents to find sheets name by numbering
    for (let i = 0; i < VANDORP_Data.length; i++) {
        let dataItem = VANDORP_Data[i];
        // find sheet by agentnumbering
        let sheetName = newWorkbook.SheetNames.find(e => e?.substring(0, 2) === dataItem[columsNames.number]?.substring(0, 2));
        if (!sheetName) continue;
        let ws = newWorkbook.Sheets[sheetName];
        if (!ws) continue;
        // target column to salary agrobox
        let colGSS = sheetColumn.gss['sheet'+(newWorkbook.SheetNames.indexOf(sheetName) + 1)];
        let colSalaryVanDorp = colGSS ? (colGSS['vandorp'] || null) : null;
        // total transport
        let important_cols = ['A', 'B'];
        let first_A_col = Object.keys(ws).find(e => e.includes(important_cols[0]));
        let line = parseInt(first_A_col.substring(1, first_A_col.length));
        const rows = xlsx.utils.sheet_to_json(ws, {header:1, blankrows: true});
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numbering = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                // get info via the WILLEMEN data by only M-CODE
                let info = VANDORP_Data.find(e => e[columsNames.mcode] === mcode && e[columsNames.number] === numbering);
                if (info) {
                    // salary 
                    if (columsNames.vandorp in info) {
                        // cols to fill
                        if (colSalaryVanDorp) {
                            let colIndex = colSalaryVanDorp+line;
                            if (!ws[colIndex]) {
                                ws[colIndex] = {t: 'n'}
                            }
                            ws[colIndex].v = info[columsNames.vandorp];
                            ws[colIndex].w = new String(info[columsNames.vandorp]);
                        }
                    }
                }
            }
            line ++;
        }
    }

    return newWorkbook;
}

// export functions
module.exports = {
    accessDB,
    // read excel file
    readWBxlsx,
    readWBxlsxstyle,
    // copy style with combination
    combineStyle,
    // save xlsx file
    saveFile,
    // get worksheet by index of the sheet
    getWS,
    // code for auth
    randomCode,
    randomnNumberCode,
    colsIndexNames,
    // others functions
    getDateNow,
    getSheetIndex,
    deleteFile,
    getVar,
    getDateInFileName,
    getFirstDateInOutputFilename,
    
    // fetch transport & répas
    arrangeTRANSPORTS,
    groupCol,
    fetchData,
    getColumnName,
    getGroupedRequiredCol,
    outTR, 
    // fetch UP
    fetchUPData,
    outUp,
    // fetch UPC data
    fetchUPCData,
    outUPC,
    // fetch agrobox data
    fetchAgroboxData,
    outAGROBOX,
    // fetch arco data
    fetchArcoData,
    outARCO,
    // fetch JeFacture data
    fetchJEFACTUREData,
    outJEFACTURE,
    // fetch PWC data
    fetchPWCData,
    outPWC,
    // fetch SPOTCHECK data
    fetchSPOTCHECKData,
    outSPOTCHECK,
    // fetch Willemen data
    fetchWillemenData,
    outWillemen,
    // fetch VanDorp data
    fetchVanDorpData,
    outVanDorp

};