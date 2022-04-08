
async function checkRequiredFiles(req, res, next) {
    console.time();
    try {
        //fs
        const fs = req.session.FS;
        // file directory
        const DIR = req.session.DIR;
        // verifier le repertoire
        if (!fs.existsSync(DIR)) {
            await fs.mkdirSync(DIR);
        }
        // files
        const FILES = await req.files;
        // file keys
        var FileKeys = await Object.keys(FILES);
        // get the global salary sheet
        const ARCOFile = await req.files['arco_salary'];
        // check files
        // NO Sheet file selected
        if (!ARCOFile) {
            return await res.send({
                status: false,
                icon: 'warning',
                message: 'No ARCO Salary file uploaded!'
            });
        }
        // No file that contains all data selected
        if (FileKeys.length === 1) {
            return await res.send({
                status: false,
                icon: 'warning',
                message: 'No ARCO Report file uploaded!'
            });
        }
        // SET VARIABLE
        req.session.ARCOFile = ARCOFile;
        req.session.FILEKEYS = FileKeys;
        req.session.FILES = FILES;
        return next();
    } catch (err) {
        console.log(err);
        return res.send({
            status: false,
            icon: 'error',
            message: 'Can not perform the program!'
        });
    }
}

async function startCorrection(req, res, next) {
    try {
        /* socket */
        await req.app.get('socket').emit('action', 'Starting correction...');
        // get passed variable
        const ARCOFile = req.session.ARCOFile;
        const DIR = req.session.DIR;
        const script = req.session.SCRIPT;
        // time to file
        const time = await new Date().getTime();
        // gs path
        const ARCOPath = await `${DIR}/${ARCOFile.name.split('.xlsx')[0]}_${time}.xlsx`;
        // COPY GSS FILE
        await ARCOFile.mv(ARCOPath);

        /* socket */
        // read sheet output file
        var wbo_sheet = await script.readWBxlsx(ARCOPath);
        await req.app.get('socket').emit('action', 'Cloning: ' + ARCOFile.name);
        
        // create the output file name
        let date = await new Date();
        const OPFileName = await `${script.getDateNow().join(".")} ARCO SALARIES WORKING CORRECTED ${date.getTime()}.xlsx`;
        const OPFilePath = await `${DIR}/${OPFileName}`;

        // SET VARIABLE
        req.session.WBO = wbo_sheet;
        req.session.OPFILENAME = OPFileName;
        req.session.OPFILEPATH = OPFilePath;
        req.session.ARCOPATH = ARCOPath;
        req.session.TIME = time;
        
        return next();
    } catch (err) {
        console.log(err);
        return res.send({
            status: false,
            icon: 'error',
            message: 'Can not perform the program!'
        });
    }
}

async function readingStyle(req, res, next) {
    try {
        const script = req.session.SCRIPT;
        /* socket */
        await req.app.get('socket').emit('action', 'Copying all style.');
        req.session.WBOS = script.readWBxlsxstyle(req.session.ARCOPATH);
        return next();
    } catch (err) {
        console.log(err);
        return res.send({
            status: false,
            icon: 'error',
            message: 'Can not perform the program!'
        });
    }
}

async function fetchAllData(req, res, next) {
    try {
        /* socket */
        await req.app.get('socket').emit('action', 'Copying arco report files.');

        // get passed variable
        const FileKeys = req.session.FILEKEYS;
        const FILES = req.session.FILES;
        const DIR = req.session.DIR;
        const script = req.session.SCRIPT;
        const wbo_sheet = req.session.WBO;
        const OPFilePath = req.session.OPFILEPATH;
        const OPFileName = req.session.OPFILENAME;
        const time = req.session.TIME;
        const sheetIndex = 0;

        // loop keys 
        await FileKeys.splice(FileKeys.indexOf('arco_salary'), 1);
        for (let i = 0; i < FileKeys.length; i++) {
            let key = await FileKeys[i];
            // switch key file
            if (key.includes('arco_report')) {
                // get file
                let file = await FILES[key];
                let filePath = await `${DIR}/${file.name.split('.xlsx')[0]}_${time}.xlsx`;

                /* socket */
                await req.app.get('socket').emit('action', 'Copying: ' + (file.name));
                // move file
                await file.mv(filePath);

                /* socket */
                await req.app.get('socket').emit('action', 'Reading: ' + (file.name));
                // read excel file
                var wbi = await script.readWBxlsxstyle(filePath);
                // set file timeout to delete
                await script.deleteFile(filePath, 25000);

                /* socket */
                await req.app.get('socket').emit('action', 'Fetching all data from: ' + (file.name));
                //.get work sheet rh
                var ws = await script.getWS(wbi, sheetIndex);
                // check sheets
                if (!ws) {
                    await req.session.WARNINGS.push({
                        status: false,
                        icon: 'warning',
                        message: `The ARCO Report file has a problem.`
                    });
                } else {
                    try {
                        /* socket */
                        await req.app.get('socket').emit('action', 'Writing all data into: ' + OPFileName);
                        // fetch all data required
                        var data = await script.getArcoCellsValue(ws, req.session.LASTINDEX - 1);
                        // set last index
                        req.session.LASTINDEX += await data.rowNumber - 1;
                        // if data is empty
                        if (Object.keys(data.cellData).length <= 0) {
                            req.session.WARNINGS.push({
                                status: false,
                                icon: 'warning',
                                message: 'No data found in the ARCO Report file number ' + (i + 1) + '.'
                            });
                        } else {
                            /* socket */
                            await req.app.get('socket').emit('action', 'Saving all from: ' + (file.name));
                            // save file
                            let output = await script.copyAndPasteARCO(data.cellData, wbo_sheet);
                            // save file
                            await script.saveFile(script.combineStyle2(output, req.session.WBOS), OPFilePath);
                        }
                    } catch (error) {
                        await console.log(error)
                        await req.session.WARNINGS.push({
                            status: false,
                            icon: 'danger',
                            message: 'There are somme errors.'
                        });
                    }
                }
            }
        }
        console.timeEnd();
        return next();
    } catch (err) {
        console.log(err);
        return res.send({
            status: false,
            icon: 'error',
            message: 'Can not perform the program!'
        });
    }
}



module.exports = {
    checkRequiredFiles,
    readingStyle,
    startCorrection,
    fetchAllData
}