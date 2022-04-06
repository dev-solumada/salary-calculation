const getWS = (wb, indexOfSheet) => {
    let sheet_name = wb.SheetNames[indexOfSheet];
    return wb.Sheets[sheet_name];
}

// get sheetname
const getSheetIndex = (wb, sheetname) => {
    return wb.SheetNames.indexOf(sheetname);
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
}

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



const combineStyle2 = (wb_xlsx, wb_xlsx_style) => {
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
                    if (cellStyled.s.fill && cellStyled.s.fill.bgColor) {
                        if (cellStyled.s.fill.fgColor.rgb === '000000') { // if bg is dark
                            cellStyled.s.fill.fgColor = {}; // set bg to white
                            let style = cellStyled.s;
                            delete style.fgColor;
                            delete cellStyled.s.fill;
                            ws_s[cellName].s = style;
                        }   
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
        while (ws[letter0+line] || ws[letter1+line]) {
            // objet pour construire un élément.
            let obj = {};
            // console.log(typeof ws[letter1+line] !== 'undefined' && typeof ws[letter1+line] !== 'undefined');
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
        let target100 = null, 
            target200 = null;
        while (line <= rows.length) {
            if (ws[important_cols[0]+line] && ws[important_cols[1]+line]) {
                // numbering agent
                let numberingagent = new String(ws[important_cols[0]+line].w).trim();
                let mcode = new String(ws[important_cols[1]+line].w).trim();
                let info = null;
                
                // PDFButler
                if (new String(numberingagent).match('PDFB-')) {
                    info = DATA_RH.find(e => e[columsNames.number] === numberingagent);
                } else if (mcode === 'GARDIEN') {
                    info = DATA_RH.find(e => e[columsNames.mcode] === 'Gardien');
                } else {
                    // get info via RH by M-CODE and Numbering Agent
                    info = DATA_RH.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                }
                if (info) {
                    // increment the numbers of agent
                    agent_found += 1;
                    var col2000 = colsToFill.find(e => e.v.v === 2000);
                    var col1000 = colsToFill.find(e => e.v.v === 1000);
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
                    
                    if (colRep)
                        if (columsNames.repas in info) {
                            if (typeof ws[colRep.c.substring(0, 1)+line] === 'undefined')
                                ws[colRep.c.substring(0, 1)+line] = {};
                            ws[colRep.c.substring(0, 1)+line].t = 'n';
                            let number = info[columsNames.mcode] === 'Gardien' ? (parseFloat(info[columsNames.repas]) || 0) : (parseFloat(info[columsNames.repas]) || 0) * 3500;
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

/**
 * SALARY EXTRACTION
 */

 const createOutputSalaryUp = (DATA_RH = [], wb, wb_style) => {
    var agent_found = 0;
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    // parcourir tous les feuilles SHEETS
    for (let i = 0; i < newWorkbook.SheetNames.length; i++) {
        let ws = newWorkbook.Sheets[newWorkbook.SheetNames[i]];
        // chercher ou se situe le 2000 et 1000
        let colSalaryUp = ''
        if (i == 0) colSalaryUp = 'G';
        if (i == 1) colSalaryUp = 'I';
        if (i == 2) colSalaryUp = 'E';
        if (i == 3) colSalaryUp = 'F';
        if (i == 4) colSalaryUp = 'E';
        
        if (i == 6) colSalaryUp = 'D';
        if (i == 8) colSalaryUp = 'D';
        
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
                        if (colSalaryUp != '') {
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

    return {
        agent_found : agent_found,
        wb: combineStyle(newWorkbook, wb_style)
    };
}

const getSalaryUPData = (ws) => {
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
            data.push(obj);
    }
    return data;
}

/**
 * AGROBOX
 */

const createOutputSalaryAGROBOX = (DATA_RH = [], wb, wb_style) => {
    var agent_found = 0;
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    // parcourir tous les feuilles SHEETS
    for (let i = 0; i < newWorkbook.SheetNames.length; i++) {
        let ws = newWorkbook.Sheets[newWorkbook.SheetNames[i]];
        // chercher ou se situe le 2000 et 1000
        let colSalaryAgrobox = ''
        if (i == 0) colSalaryAgrobox = 'F';
        
        if (i == 2) colSalaryAgrobox = 'G';
        if (i == 3) colSalaryAgrobox = 'D';
        if (i == 4) colSalaryAgrobox = 'H';
        
        if (i == 6 || i == 8) colSalaryAgrobox = 'C';
        
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
                        if (colSalaryAgrobox != '') {
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

    return {
        agent_found : agent_found,
        wb: combineStyle(newWorkbook, wb_style)
    };
    
}

const getSalaryAgroboxData = (ws) => {
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
            data.push(obj);
    }
    return data;
}

/**
 * ACRO
 */

const createOutputSalaryARCO = (DATA_RH = [], wb, wb_style) => {
    var agent_found = 0;
    const xlsx = require('xlsx');
    // creer un nouveau work book
    var newWorkbook = wb;
    // parcourir tous les feuilles SHEETS
    for (let i = 0; i < newWorkbook.SheetNames.length; i++) {
        let ws = newWorkbook.Sheets[newWorkbook.SheetNames[i]];
        // chercher ou se situe le 2000 et 1000
        let colSalaryArco = ''
        if (i == 0 || i == 5) colSalaryArco = 'D';
        
        if (i == 1 || i == 2) colSalaryArco = 'C';
        
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
                let info = DATA_RH.find(e => e[columsNames.number] === numberingagent && e[columsNames.mcode] === mcode);
                if (info) {
                    // salary 
                    if (columsNames.salaryARCO in info) {
                        // cols to fill
                        if (colSalaryArco != '') {
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

    return {
        agent_found : agent_found,
        wb: combineStyle(newWorkbook, wb_style)
    };
    
}

const getSalaryArcoData = (ws) => {
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
                    obj[columsNames.salaryARCO] = parseFloat(cell.v) || 0;
                }
            }
        }
        // if keys exist
        if ((columsNames.mcode in obj || columsNames.number in obj) && columsNames.salaryARCO in obj) 
            data.push(obj);
    }
    return data;
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
 * CORRECT ARCO SALARY
 */
// 25 - 1907
const getArcoCellsValue = (ws, lastIndex) => {
    const XLSX = require('xlsx');
    var range = XLSX.utils.decode_range(ws['!ref']);
    var data = {};
    // rows
    for (let rowNum = 24; rowNum <= range.e.r; rowNum++) {
        // cells
        if (typeof ws['A'+rowNum] === 'object' && ws['A'+rowNum].w !== undefined) {
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
                let cell = ws[cellName];
                if (cell) {
                    let newCell = cellName.substring(0, 1) + parseInt((cellName.substring(1, cellName.length) - 23 + lastIndex));
                    data[newCell] = cell;
                }
            }
        }
    }
    return data;
}


const copyAndPasteARCO = (data, wb) => {
    var newWorkbook = setFormula(wb);
    let ws = newWorkbook.Sheets[newWorkbook.SheetNames[0]];
    Object.keys(data).forEach(key => {ws[key] = data[key];});

    // const XLSX = require('xlsx');
    // var range = XLSX.utils.decode_range(ws['!ref']);
    // data1 = getARCOValidationFiltered(ws);
    // let v = getTotalValidationARCO(data1, 'M-KTN');
    // for (let rowNum = 24; rowNum <= range.e.r; rowNum++) {
    //     let cell = ws['I'+rowNum];
    //     if (typeof cell === 'object') {
    //         let v = getTotalValidationARCO(data, cell.w || '');
    //         if (v) {
    //             ws['D'+rowNum].v = v.v1Total;
    //             ws['D'+rowNum].w = new String(v.v1Total);
    //         } 
    //     }
    // }

    return newWorkbook;
}

const setFormula = (wb) => {
    const XLSX = require('xlsx');
    let ws = wb.Sheets[wb.SheetNames[1]];
    // looping throup sheet 2
    var range = XLSX.utils.decode_range(ws['!ref']);
    // rows
    for (let rowNum = 9; rowNum <= range.e.r; rowNum++) {
        let f = `SUMIF(${wb.SheetNames[0]}!C2:'${wb.SheetNames[0]}'!C${range.e.r},Summary!C${rowNum},${wb.SheetNames[0]}!D2:'${wb.SheetNames[0]}'!D${range.e.r})`;
        if (typeof ws['D'+rowNum] === 'object') ws['D'+rowNum].f = f;
        f = `SUMIF(${wb.SheetNames[0]}!F2:'${wb.SheetNames[0]}'!F${range.e.r},Summary!C${rowNum},${wb.SheetNames[0]}!D2:'${wb.SheetNames[0]}'!D${range.e.r})`;
        if (typeof ws['E'+rowNum] === 'object') ws['E'+rowNum].f = f;
    }
    return wb;
}

const getARCOValidationFiltered = (ws) => {
    const XLSX = require('xlsx');
    var range = XLSX.utils.decode_range(ws['!ref']);
    var data = [];
    // rows
    for (let rowNum = 24; rowNum <= range.e.r; rowNum++) {
        // cells
        let obj = {};
        for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
            let cellName = XLSX.utils.encode_cell({r: rowNum, c: colNum});
            let cell = ws[cellName];
            if (cell) {
                if (cellName.includes('A')) obj[ws['A24'].v] = cell.w;
                if (cellName.includes('C')) obj['M-CODE'] = cell.w;
                if (cellName.includes('D')) obj['v1'] = cell.v;
                if (cellName.includes('E')) obj['v2'] = cell.v;
            }
        }
        if (Object.keys(obj).length > 0) data.push(obj);
    }
    return data;
}

const getTotalValidationARCO = (data = [], filterKey) => {
    let newData = data.filter(e => e['M-CODE'] === filterKey);
    if (newData.length > 0) {
        let v1Total = newData.map(e => e['v1']).reduce((a, b) => a + b, 0);
        let v2Total = newData.map(e => e['v2']).reduce((a, b) => a + b, 0);
        return {'v1Total': v1Total, 'v2Total': v2Total};
    } else {
        return null;
    }
}

// export functions
module.exports = {
    readWBxlsx,
    readWBxlsxstyle,
    combineStyle,
    combineStyle2,
    arrangeTRANSPORTS,
    getColumnName,
    groupCol,
    fetchData,
    createOutput,
    createOutputSalaryUp,
    saveFile,
    getWS,
    getGroupedRequiredCol,
    randomCode,
    randomnNumberCode,
    colsIndexNames,
    getSalaryUPData,
    getSalaryAgroboxData,
    createOutputSalaryAGROBOX,
    getSalaryArcoData,
    createOutputSalaryARCO,
    getDateNow,
    getSheetIndex,
    deleteFile,
    // correction
    getArcoCellsValue,
    copyAndPasteARCO,
    getARCOValidationFiltered,
    getTotalValidationARCO,

    setFormula
};