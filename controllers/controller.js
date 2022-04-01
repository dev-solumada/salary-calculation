require('dotenv').config();
const express = require('express');
const router = express.Router();
const fs = require('fs');
const script = require('./script.js')
const mongoose = require('mongoose');
const UserSchema = require('../models/UserSchema');
const SCSchema = require('../models/SCSchema');
const NotifSchema = require('../models/NotifSchema');
const nodemailer = require('nodemailer');
const moment = require('moment');
var currentFile = null, finished = false;
//Mailing
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.MAIN_USER,
    pass: process.env.MAIN_PASS,
  },
  tls: {
    rejectUnauthorized: false
  }
});
// mongoo options
const MongooOptions = {
    useUnifiedTopology: true,
    UseNewUrlParser: true,
};

const getAllUsers = async () => {
    let user = [];
    await mongoose.connect(
    process.env.MONGO_URI,
    {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
    }
    ).then(async () => {
        user = await UserSchema.find();
    }).catch(err => { });

    return user;
}

const getSCInfo = async () => {
    let user = [];
    await mongoose.connect(
    process.env.MONGO_URI,
    {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
    }
    ).then(async () => {
        user = await SCSchema.find();
    }).catch(err => { });

    return user[0];
}

// get notifications
const getNotifs = async () => {
    const Limit = 8;
    let notifs = [];
    await mongoose.connect(
    process.env.MONGO_URI,
    {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
    }
    ).then(async () => {
        allNotifs = await NotifSchema.find();
        notifs = await NotifSchema.find().sort('-creation').limit(Limit);
        notifs.forEach(async e => {e.moment = moment(e.creation).fromNow()});
        if (allNotifs.length > Limit) {
            // supprimer les autres notifications
            await NotifSchema.deleteMany();
            // insert many 
            await NotifSchema.insertMany(notifs);
        }
        notifs.forEach(n => {
            if (n.category === 'salary calculation' || n.category === 'correct arco') {
                n.file = n.link;
                n.exists = true;
                if (!fs.existsSync('uploads/'+n.link)) {
                    n.exists = false;
                    n.link = '#';
                }
            }
        })
    }).catch(err => { });

    return notifs;
}

// redirect to login
const redirectLogin = (req, res, next) => {
    if (req.session.userId) {
        if (!req.session.userId.access)
            return (res.redirect('/login'));
        return next();
    } else {
        return res.redirect('/login');
    }
}

const checkType = (req, res, next) => {
    if (req.session.userId.usertype !== 'admin') {
        backURL=req.header('Referer') || '/';
        res.redirect(backURL);
    } else {
        next();
    }
}

// redirect to home
const redirectHome = (req, res, next) => {
    if (!req.session.userId) {
        next();
    } else {
        res.redirect('/home');
    }
}

// Navigation 
router.route('').get((req, res) => {
    return res.redirect('/home');
});

// LOGOUT
router.route('/logout').get(redirectLogin, (req, res) => {
    req.session.destroy();
    
    return res.redirect('/login');
});

// FORGOT
router.route('/forgot-password').get(redirectHome, (req, res) => {
    req.session.foundUser = null;
    return res.render('forgot', {login: true, year: new Date().getFullYear()});
});

// CODE
router.route('/enter-code').all(redirectHome, (req, res) => {
    const codeAuth = req.session.code;
    if (req.method === 'GET') {
        if (codeAuth) 
            return res.render('enter-code', {login: true, year: new Date().getFullYear()});
        else 
            return res.redirect('/forgot-password');
    } else {
        const {code} = req.body;
        if (code === codeAuth) {
            res.send({status: true});
        } else {
            res.send({status: false, message: 'You have entered an invalid code.'})
        }
    }
});

// NEW PASSWORD & SET A NEW PASSWORD
router.route('/new-password').all(redirectHome, (req, res) => {
    const foundUser = req.session.foundUser;
    if (req.method === 'GET') {
        if (foundUser)
            res.render('new-password', {login: true, year: new Date().getFullYear()});
        else 
            return res.redirect('/forgot-password');
    } else {
        mongoose.connect(
            process.env.MONGO_URI,
            {
                useUnifiedTopology: true,
                UseNewUrlParser: true,
            }
        ).then(async () => {
            const {password, keep} = req.body;
            foundUser.password = password;
            await UserSchema.findOneAndUpdate({email: foundUser.email}, foundUser);
            // if user checked the keep sign in
            if (keep === 'true') {
                req.session.userId = foundUser;
            }
            res.send({status: true});
        }).catch(err => {
            return res.send({
                status: false,
                target: 'database',
                message: err
            });
        });
    }
});


// CHECK USER NAME
router.route('/find-username').post(redirectHome, (req, res) => {
    const {username} = req.body;
    
    mongoose.connect(
        process.env.MONGO_URI,
        MongooOptions
    ).then(async () => {
        const result = await UserSchema.findOne({$or: [{username: username}, {email: username}]});
        // if result is not null
        if (result) {
            // generate random number code
            let random_code = script.randomnNumberCode();
            // set session number code
            req.session.code = random_code;
            // send email
            var mailOptions = {
                from: process.env.MAIN_USER,
                to: result.email,
                subject: "Validation code for Salary Calculation",
                html:
                    `<div style="padding: 8px; background-color: aliceblue;">
                    <center>
                    <h1>Salary Calculation Validation Code</h1>
                    <table border="1" style="background:white; width: 400px; border-color: #eee; padding: 8px;border-collapse: collapse;">
                        <tbody>
                            <tr>
                                <th style="text-align: left;  padding: 8px !important;">CODE:</th>
                                <td style="padding: 8px !important;">${random_code}</td>
                            </tr>
                        </tbody>
                    </table>
                    <br>
                    </center>
                </div>`,
            };
            await transporter.sendMail(mailOptions, function (error, info) {
                error ? console.log(error) : console.log(info);
                transporter.close();
            });
            req.session.foundUser = result;
            res.send({status: true})
        } else {
            res.send({
                status: false,
                target: 'username',
                message: 'This username or email is not found.'
            });
        }

    }).catch(err => {
        console.log(object);
        return res.send({
            status: false,
            target: 'database',
            message: err
        });
    });
});


// LOGIN 
router.route('/login').all(redirectHome, (req, res) => {
    if (req.method === 'GET') {
        return res.render('login', {login: true, year: new Date().getFullYear()});
    } else {
        mongoose.connect(
            process.env.MONGO_URI,
            {
                useUnifiedTopology: true,
                UseNewUrlParser: true,
            }
        ).then(async () => {
            let user = await {
                username: req.body.username,
                password: req.body.password,
            }
    
            // check username
            let find_user = await UserSchema.findOne({username: user.username});

            // if username found
            let data = [];
            if (find_user) {
                if (user.password !== find_user.password)
                data.push({
                    target: 'password',
                    message: 'Your password is wrong'
                });
                // check access
                if (!find_user.access) {
                    data.push({
                        target: 'access',
                        message: 'Sorry, your access is closed. Please, contact the admin.'
                    });
                }
            } else {
                data.push({
                    target: 'username',
                    message: 'The username doesn\'t exist.'
                });
            } 
    
            // if data is not empty
            if (data.length > 0) {
                res.send({
                    status: false,
                    data: data
                });
            } else {
                req.session.userId = find_user;
                // go to home
                res.send({
                    status: true,
                    message: 'Login with success!'
                });

                // set notif
                let notif = {
                    category: 'user',
                    description: 'A user was logging in',
                    creation: new Date()
                }
                await new NotifSchema(notif).save();
            }
        }).catch(err => {
            console.log(err);
            res.send({
                target: 'database',
                status: false,
                message: 'Unable to connect the database.'
            });
        });
    }
});

/**
 * HOME Page.
 */
router.route('/home').get(redirectLogin, async (req, res) => {
    const user = req.session.userId;
    const allUsers = await getAllUsers();
    // last added user
    let lastAddUser = moment(allUsers[allUsers.length - 1].creation).fromNow();
    let date = new Date();
    // last action on SC
    let scInfo = await getSCInfo();
    let lastSCInfo = (scInfo) ? moment(scInfo.creation).fromNow() : 'none';
    let csInfoActivity = (scInfo) ? scInfo.number : 0;
    // notifications
    let notifs = await getNotifs();
    res.render('index', {
        login: false,
        active: 'home',
        active_sub: '',
        year: date.getFullYear(),
        date_short: script.getDateNow().join("/"),
        date_long: date.toDateString(),
        user: user,
        allUsers: allUsers,
        lastAddUser: lastAddUser,
        lastSCInfo: lastSCInfo,
        csInfoActivity: csInfoActivity,
        notifs: notifs
    });
});

/**
 * PROGRAMME - Salary Calculation.
 */
router.route('/salary-calculation').get(redirectLogin, async (req, res) => {
    const user = req.session.userId;
    // notifications
    let notifs = await getNotifs();
    res.render('salary-calculation', {login: false, active: 'programmes', active_sub: 'salary-calculation', year: new Date().getFullYear(), user: user, notifs: notifs});
});

/**
 * USERS - Add new user.
 */
router.route('/add-new-user').get(redirectLogin, async (req, res) => {
    const user = req.session.userId;
    // notifications
    let notifs = await NotifSchema.find();
    notifs.forEach(async e => {e.moment = moment(e.creation).fromNow()});
    res.render('add-new-user', {login: false, active: 'users', active_sub: 'add-new-user', year: new Date().getFullYear(), user: user, notifs: notifs});
});

/**
 * USERS - set user access.
 */
router.route('/set-user-access').post(redirectLogin, (req, res) => {
    mongoose.connect(
    process.env.MONGO_URI,
    MongooOptions
    ).then(async () => {
        // find all users
        let user = {access: req.body.access}
        await UserSchema.findOneAndUpdate({email: req.body.email}, user);
        // set notif
        let notif = {
            category: 'user',
            description: 'An user access has been ' + (req.body.access == 'true' ? 'given' : 'closed'),
            creation: new Date()
        }
        await new NotifSchema(notif).save();

        await res.send({
            status: true,
            message: 'Access successfully set.'
        });
    }).catch(err => {
        res.send({
            target: 'database',
            status: false,
            message: 'Unable to connect the database.'
        });
    })
});

/**
 * USERS - Delete new user.
 */
router.route('/delete-user').post(redirectLogin, (req, res) => {
    mongoose.connect(
        process.env.MONGO_URI,
        MongooOptions
    ).then(async () => {
        // find all users
        await UserSchema.findOneAndDelete({email: req.body.email});
        // set notif
        let notif = {
            category: 'user',
            description: 'An user has been deleted.',
            creation: new Date()
        }
        await new NotifSchema(notif).save();
        // send response
        res.send({
            status: true,
            message: 'User successfully removed.'
        });
    }).catch(err => {
        res.send({
            target: 'database',
            status: false,
            message: 'Unable to connect the database.'
        });
    })
});

/**
 * USERS - Users list.
 */
router.route('/users-list').get(redirectLogin, checkType, async (req, res) => {
    const user = req.session.userId;
    mongoose.connect(
        process.env.MONGO_URI,
        MongooOptions
    ).then(async () => {
        // find all users
        let users = await UserSchema.find();
        // notifications
        let notifs = await getNotifs();
        return res.render('users-list', {login: false, active: 'users', active_sub: 'users-list', year: new Date().getFullYear(), users: users, user: user, notifs: notifs});
    }).catch(err => {
        res.send({
            target: 'database',
            status: false,
            message: 'Unable to connect the database.'
        });
    })
});


/**
 * USERS - Manage access.
 */
 router.route('/manage-access').get(redirectLogin, checkType, async (req, res) => {
    const user = req.session.userId;
    mongoose.connect(
        process.env.MONGO_URI,
        MongooOptions
    ).then(async () => {
        // find all users
        let users = await UserSchema.find();
        // notifications
        let notifs = await getNotifs();
        res.render('manage-access', {login: false, active: 'users', active_sub: 'manage-access', year: new Date().getFullYear(), users: users, user: user, notifs: notifs});
    }).catch(err => {
        res.send({
            target: 'database',
            status: false,
            message: 'Unable to connect the database.'
        });
    })
});

/**
 * USERS - Edit user.
 */
 router.route('/edit-user/:email').all(redirectLogin, checkType, async (req, res) => {
    const userS = req.session.userId;
    mongoose.connect(
        process.env.MONGO_URI,
        MongooOptions
    ).then(async () => {
        if (req.method === 'POST') {
            let user = await {
                username: req.body.username,
                email: req.body.email,
            }
            // check username
            let userSelected = await UserSchema.findOne({email: user.email});
            let find_username = await UserSchema.findOne({username: user.username});
            let find_email = await UserSchema.findOne({email: user.email});
            // if username found
            let data = [];
            if (find_username) {
                if (userSelected.username !== user.username)
                    data.push({
                        target: 'username',
                        message: 'The username is already taken. Please, try another.'
                    });  
            } 
            // if username found
            if (find_email) {
                console.log(userSelected.email, req.params.email);
                if (userSelected.email !== req.params.email)
                    data.push({
                        target: 'email',
                        message: 'The email address is already used. Please, try another.'
                    });
            } 
    
            // if data is not empty
            if (data.length > 0) {
                res.send({
                    status: false,
                    data: data
                });
            } else {
                // save user
                await UserSchema.findOneAndUpdate({email: req.params.email}, user);

                // check if active user
                if (req.params.email === userS.email) {
                    let token = await UserSchema.findOne({email: user.email});
                    req.session.userId = token;
                }
                
                // set notif
                let notif = {
                    category: 'user',
                    description: 'An user has been updated.',
                    creation: new Date()
                }
                await new NotifSchema(notif).save();
                
                await res.send({
                    status: true,
                    message: 'User successfully updated!'
                });
            }
        } else { // GET
            // find all users
            let userEdit = await UserSchema.findOne({email: req.params.email});
            if (userEdit) {
                // notifications
                let notifs = await NotifSchema.find();
                notifs.map(async e => e.creation = moment(e.creation).fromNow());
                return res.render('edit-user', {login: false, active: 'users', active_sub: 'users-list', year: new Date().getFullYear(), userEdit: userEdit, user: userS, notifs: notifs});
            } else {
                res.redirect('/users-list');
            }
        }
    }).catch(err => {
        console.log(err);
        res.send({
            target: 'database',
            status: false,
            message: 'Unable to connect the database.'
        });
    })
});

/**
 * Add user
 */
router.route('/add-user').post(redirectLogin, checkType, (req, res) => {
    mongoose.connect(
        process.env.MONGO_URI,
        MongooOptions
    ).then(async () => {
        let user = await req.body;

        let random_code = req.body.random_code;
        // check username
        let find_username = await UserSchema.findOne({username: user.username});
        let find_email = await UserSchema.findOne({email: user.email});
        // if username found
        let data = [];
        if (find_username) {
            data.push({
                target: 'username',
                message: 'The username is already taken. Please, try another.'
            });  
        } 
        // if username found
        if (find_email) {
            data.push({
                target: 'email',
                message: 'The email address is already used. Please, try another.'
            });
        } 

        // if data is not empty
        if (data.length > 0) {
            res.send({
                status: false,
                data: data
            });
        } else {
            if (random_code) {
                // generate random code
                let random = script.randomCode();
                user.password = random;
                var mailOptions = {
                    from: process.env.MAIN_USER,
                    to: user.email,
                    subject: "Authentification code for Salary Calculation",
                    html:
                      `<div style="padding: 8px; background-color: aliceblue;">
                      <center>
                      <h1>Salary Calculation Authentification</h1>
                      <table border="1" style="background:white; width: 400px; border-color: #eee; padding: 8px;border-collapse: collapse;">
                          <tbody>
                              <tr>
                                  <th style="text-align: left;  padding: 8px !important;">Your username:</th>
                                  <td style="padding: 8px !important;">${user.username}</td>
                              </tr>
                              <tr>
                                  <th style="text-align: left;  padding: 8px !important;">Your password:</th>
                                  <td style="padding: 8px !important;">${user.password}</td>
                              </tr>
                          </tbody>
                      </table>
                      <br>
                      </center>
                    </div>`,
                };
                await transporter.sendMail(mailOptions, function (error, info) {
                    error ? console.log(error) : console.log(info);
                    transporter.close();
                });
                
            }
                
            // set notif
            let notif = {
                category: 'user',
                description: 'An user has been added.',
                creation: new Date()
            }
            await new NotifSchema(notif).save();
        
            // save user
            await UserSchema(user).save();
            await res.send({
                status: true,
                message: 'Users successfully registrered!'
            });
        }
    }).catch(err => {
        console.log(err);
        res.send({
            target: 'database',
            status: false,
            message: 'Unable to connect the database.'
        });
    });
})




/**
 * UPLOAD - RH file and Salary Sheet file.
 */
 router.route('/upload-xlsx').post(redirectLogin, async (req, res) => { 
    try {
        if(!req.files) {
            return res.send({
                status: false,
                icon: 'warning',
                message: 'No files uploaded!'
            });
        } else {
            // file directory
            const DIR = await 'uploads';
            // verifier le repertoire
            if (!fs.existsSync(DIR)) {
                await fs.mkdirSync(DIR);
            }
            // files
            const FILES = req.files;
            // file keys
            const FileKeys = await Object.keys(FILES);
            // get the global salary sheet
            const GSSFile = await req.files['sheet_file'];
            // check files
            // NO Sheet file selected
            if (!GSSFile) {
                return res.send({
                    status: false,
                    icon: 'warning',
                    message: 'No GLOBAL SALARY Sheet file uploaded!'
                });
            }
            // No file that contains all data selected
            if (FileKeys.length === 1) {
                return res.send({
                    status: false,
                    icon: 'warning',
                    message: 'No file that contains data uploaded!'
                });
            } 
            // time to file
            const time = new Date().getTime();
            // gs path
            const GSSPATH = await `${DIR}/${GSSFile.name.split('.xlsx')[0]}_${time}.xlsx`;
            // COPY GSS FILE
            await GSSFile.mv(GSSPATH);
            // read sheet output file
            var wbo_sheet = await script.readWBxlsx(GSSPATH);
            var wbo_sheet_style = await script.readWBxlsxstyle(GSSPATH);
            // create the output file name
            let date = new Date();
            const OPFileName = await `${script.getDateNow().join(".")} GSS ${date.getTime()}.xlsx`;
            const OPFilePath = await `${DIR}/${OPFileName}`;
            // warnigngs
            const Warnings = await [];
            // step
            var step = 0;
            // loop keys 
            for (let i = 0; i < FileKeys.length; i++) {
                let key = FileKeys[i];
                // get file
                let file = await FILES[key];
                let filePath = await `${DIR}/${file.name.split('.xlsx')[0]}_${time}.xlsx`;
                // move file
                await file.mv(filePath);
                // set file timeout to delete
                await script.deleteFile(filePath, 30000);
                // read excel file
                var wbi = await script.readWBxlsx(filePath);
                // switch key file
                switch (key) {
                    case 'rh_file':
                        var sheetName = 'Repas & Transport';
                        var sheetIndex = script.getSheetIndex(wbi, sheetName);
                        //.get work sheet rh
                        var ws = await script.getWS(wbi, sheetIndex);
                        // check sheets
                        if (!ws) {
                            Warnings.push({
                                status: false,
                                icon: 'warning',
                                message: `The RH file has a problem. No "${sheetName}" sheetname found. Please check the file.`
                            });
                        } else {
                            try {
                                // fetch all data required
                                var data = await script.fetchData(ws);
                                // if data is empty
                                if (data.length <= 0) {
                                    Warnings.push({
                                        status: false,
                                        icon: 'warning',
                                        message: 'No data found in the RH file.'
                                    });
                                } else {
                                    // output file
                                    let output = await script.createOutput(data, wbo_sheet, wbo_sheet_style);
                                    if (output.agent_found === 0) {
                                        Warnings.push({
                                            status: false,
                                            icon: 'warning',
                                            message: 'No Agent and Required Columns found in the GLOBAL SALARY SHEET! Please verify the file.'
                                        });
                                    } else {
                                        // if step one is done change the to the output file.
                                        if (step !== 0)
                                        wbo_sheet = await script.readWBxlsx(OPFilePath);
                                        let output = await script.createOutput(data, wbo_sheet, wbo_sheet_style);
                                        // save file
                                        await script.saveFile(output.wb, OPFilePath);
                                        step = await step + 1;
                                    }
                                }
                            } catch (error) {
                                Warnings.push({
                                    status: false,
                                    icon: 'danger',
                                    message: 'The RH file has a big problem.'
                                });
                            }
                        }
                        break;
                    // UNIFIED POST
                    case 'salaryup_file':
                        var sheetName = 'UnifiedPost Salaris per agent';
                        var sheetIndex = script.getSheetIndex(wbi, sheetName);
                        ws = await script.getWS(wbi, sheetIndex);
                        if (!ws) {
                            Warnings.push({
                                status: false,
                                icon: 'warning',
                                message: `The UP Salary file has a problem. No "${sheetName}" sheetname found. Please verify the file.`
                            });
                        } else {
                            try {
                                data = await script.getSalaryUPData(ws);
                                // if data is empty
                                if (data.length <= 0) {
                                    Warnings.push({
                                        status: false,
                                        icon: 'warning',
                                        message: 'No data found in the UP Salary file! Please verify it.'
                                    });
                                } else {
                                    // if step one is done change the to the output file.
                                    if (step !== 0)
                                        wbo_sheet = await script.readWBxlsx(OPFilePath);
                                    let output = await script.createOutputSalaryUp(data, wbo_sheet, wbo_sheet_style);
                                    await script.saveFile(output.wb, OPFilePath);
                                    step = await step + 1;
                                }
                            } catch (error) {
                                Warnings.push({
                                    status: false,
                                    icon: 'danger',
                                    message: 'The UP Salary file has a big problem.'
                                });
                            }
                        }
                        break;
                    // AGROBOX
                    case 'salaryagrobox_file':
                        var sheetName = 'agrobox salaries per agent';
                        var sheetIndex = script.getSheetIndex(wbi, sheetName);
                        ws = await script.getWS(wbi, sheetIndex);
                        if (!ws) {
                            Warnings.push({
                                status: false,
                                icon: 'warning',
                                message: `The Agrobox Salary file has a problem. No "${sheetName}" sheetname found. Please verify the file.`
                            });
                        } else {
                            try {
                                data = await script.getSalaryAgroboxData(ws);
                                // if data is empty
                                if (data.length <= 0) {
                                    Warnings.push({
                                        status: false,
                                        icon: 'warning',
                                        message: 'No data found in the Agrobox Salary file! Please verify it.'
                                    });
                                } else {
                                    // if step one is done change the to the output file.
                                    if (step !== 0)
                                        wbo_sheet = await script.readWBxlsx(OPFilePath);
                                    let output = await script.createOutputSalaryAGROBOX(data, wbo_sheet, wbo_sheet_style);
                                    await script.saveFile(output.wb, OPFilePath);
                                    step = await step + 1;
                                }
                            } catch (error) {
                                Warnings.push({
                                    status: false,
                                    icon: 'danger',
                                    message: 'The Agrobox Salary file has a big problem.'
                                });
                            }
                        }
                        break;
                    case 'salaryarco_file':
                        var sheetName = 'Summary';
                        var sheetIndex = script.getSheetIndex(wbi, sheetName);
                        ws = await script.getWS(wbi, 1);

                        if (!ws) {
                            Warnings.push({
                                status: false,
                                icon: 'warning',
                                message: `The Arco Salary file has a problem. No "${sheetName}" sheetname found. Please verify the file.`
                            });
                        } else {
                            try {
                                data = await script.getSalaryArcoData(ws);
                                // if data is empty
                                if (data.length <= 0) {
                                    Warnings.push({
                                        status: false,
                                        icon: 'warning',
                                        message: 'No data found in the Arco Salary file! Please verify it.'
                                    });
                                } else {
                                    // if step one is done change the to the output file.
                                    if (step !== 0)
                                        wbo_sheet = await script.readWBxlsx(OPFilePath);
                                    let output = await script.createOutputSalaryARCO(data, wbo_sheet, wbo_sheet_style);
                                    await script.saveFile(output.wb, OPFilePath);
                                    step = await step + 1;
                                }
                            } catch (error) {
                                console.log(error);
                                Warnings.push({
                                    status: false,
                                    icon: 'danger',
                                    message: 'The Agrobox Salary file has a big problem.'
                                });
                            }
                        }
                        break;
                    default:
                        break;
                }
            }

            // FINISHED check file
            if (step > 0 && fs.existsSync(OPFilePath)) {
                // set timeout for the output file
                await setTimeout(() => {
                    fs.unlinkSync(OPFilePath);
                }, 1000 * 60 * 60);
                //send response
                await res.send({
                    status: true,
                    icon: 'success',
                    message: 'The file is proccessed successfully.',
                    file: OPFileName,
                    warnings: Warnings
                });
                finished = true;
                // save info to database
                mongoose.connect(
                    process.env.MONGO_URI,
                    {
                        useUnifiedTopology: true,
                        UseNewUrlParser: true,
                    }
                ).then(async () => {
                    let info = await SCSchema.find();
                    if (info[0]) {
                        info[0].number += 1;
                        info[0].creation = new Date();
                        await SCSchema.findOneAndUpdate({name: 'info'}, info[0])
                    } else {
                        let newinfo = {
                            name: 'info',
                            number: 1
                        }
                        await new SCSchema(newinfo).save();
                    }
                    
                    // set notif
                    let notif = {
                        category: 'salary calculation',
                        description: 'Salary calculation: recent activity',
                        creation: new Date(),
                        link: OPFileName,
                        user: req.session.userId.username
                    }
                    await new NotifSchema(notif).save();
                    
                }).catch(err => {
                    res.send({
                        target: 'database',
                        status: false,
                        message: 'Unable to connect the database.'
                    });
                });
            } else {
                //send response
                await res.send({
                    status: false,
                    icon: 'warning',
                    message: 'Can not perform the program.',
                    file: OPFileName,
                    warnings: Warnings
                });
            }
            return;
        }
    } catch (err) {
        console.log(err)
        res.status(500).send({status: false, icon: 'error', message: 'Server error!'});
    }
});



/**
 * UPLOAD - ARCO FILE
 */
router.route('/upload-correct-arco').post(redirectLogin, async (req, res) => { 
    try {
        if(!req.files) {
            return res.send({
                status: false,
                icon: 'warning',
                message: 'No files uploaded!'
            });
        } else {
            // file directory
            const DIR = await 'uploads';
            // verifier le repertoire
            if (!fs.existsSync(DIR)) {
                await fs.mkdirSync(DIR);
            }
            // files
            const FILES = req.files;
            // file keys
            const FileKeys = await Object.keys(FILES);
            // get the global salary sheet
            const ARCOFile = await req.files['arco_salary'];
            // check files
            // NO Sheet file selected
            if (!ARCOFile) {
                return res.send({
                    status: false,
                    icon: 'warning',
                    message: 'No ARCO Salary file uploaded!'
                });
            }
            // No file that contains all data selected
            if (FileKeys.length === 1) {
                return res.send({
                    status: false,
                    icon: 'warning',
                    message: 'No ARCO Report file uploaded!'
                });
            } 
            // time to file
            const time = new Date().getTime();
            // gs path
            const ARCOPath = await `${DIR}/${ARCOFile.name.split('.xlsx')[0]}_${time}.xlsx`;
            // COPY GSS FILE
            await ARCOFile.mv(ARCOPath);
            // read sheet output file
            var wbo_sheet = await script.readWBxlsx(ARCOPath);
            var wbo_sheet_style = await script.readWBxlsxstyle(ARCOPath);
            // create the output file name
            let date = new Date();
            const OPFileName = await `${script.getDateNow().join(".")} ARCO SALARIES WORKING CORRECTED ${date.getTime()}.xlsx`;
            const OPFilePath = await `${DIR}/${OPFileName}`;
            // set file name in a session
            currentFile = OPFileName;
            // warnigngs
            const Warnings = await [];
            // data from acro report
            let lastIndex = 0;
            // loop keys 
            for (let i = 0; i < FileKeys.length; i++) {
                let key = FileKeys[i];
                // get file
                let file = await FILES[key];
                let filePath = await `${DIR}/${file.name.split('.xlsx')[0]}_${time}.xlsx`;
                // move file
                await file.mv(filePath);
                // set file timeout to delete
                await script.deleteFile(filePath, 30000);
                // read excel file
                var wbi = await script.readWBxlsxstyle(filePath);
                // switch key file
                if (key.includes('arco_report')) {
                    var sheetIndex = 0 ; //script.getSheetIndex(wbi, sheetName);
                    //.get work sheet rh
                    var ws = await script.getWS(wbi, sheetIndex);
                    // check sheets
                    if (!ws) {
                        Warnings.push({
                            status: false,
                            icon: 'warning',
                            message: `The ARCO Report file has a problem.`
                        });
                    } else {
                        try {
                            // fetch all data required
                            var data = await script.getArcoCellsValue(ws, lastIndex);
                            // if data is empty
                            if (Object.keys(data).length <= 0) {
                                Warnings.push({
                                    status: false,
                                    icon: 'warning',
                                    message: 'No data found in the ARCO Report file number ' + (i + 1) + '.'
                                });
                            } else {

                                // save file
                                let output = await script.combineStyle2(script.copyAndPasteARCO(data, wbo_sheet), wbo_sheet_style);
                                // save file
                                await script.saveFile(output, OPFilePath);
                                // set last index
                                let cellName = await Object.keys(data)[Object.keys(data).length-1];
                                lastIndex = await parseInt(cellName.substring(1, cellName.length)); 
                            }
                        } catch (error) {
                            console.log(error)
                            Warnings.push({
                                status: false,
                                icon: 'danger',
                                message: 'There are somme errors.'
                            });
                        }
                    }
                }
            }

            // FINISHED check file
            if (fs.existsSync(OPFilePath)) {
                //send response
                await res.send({
                    status: true,
                    icon: 'success',
                    message: 'The file is proccessed successfully.',
                    file: OPFileName,
                    warnings: Warnings
                });
                // set timeout for the output file
                await setTimeout(() => {
                    fs.unlinkSync(OPFilePath);
                }, 1000 * 60 * 60);
                console.log('vita');
                // save info to database
                mongoose.connect(
                    process.env.MONGO_URI,
                    {
                        useUnifiedTopology: true,
                        UseNewUrlParser: true,
                    }
                ).then(async () => {
                    // set notif
                    let notif = {
                        category: 'correct arco',
                        description: 'ARCO Correction: Recent Activity',
                        creation: new Date(),
                        link: OPFileName,
                        user: req.session.userId.username
                    }
                    await new NotifSchema(notif).save();
                    
                }).catch(err => {
                    res.send({
                        target: 'database',
                        status: false,
                        message: 'Unable to connect the database.'
                    });
                });
            } else {
                //send response
                await res.send({
                    status: false,
                    icon: 'warning',
                    message: 'Can not perform the program.',
                    file: OPFileName,
                    warnings: Warnings
                });
            }
            return;
        }
    } catch (err) {
        console.log(err)
        res.status(500).send({status: false, icon: 'error', message: 'Server error!'});
    }
});

router.route('/downloading').post(redirectLogin, checkType, async (req, res) => {
    res.send({
        status: true,
        icon: 'success',
        message: 'The file is proccessed successfully.',
        file: currentFile,
    });
    
});

router.route('/correct-arco').get(redirectLogin, checkType, async (req, res) => {
    
    const user = req.session.userId;
    mongoose.connect(
        process.env.MONGO_URI,
        MongooOptions
    ).then(async () => {
        // notifications
        let notifs = await getNotifs();
        return res.render('correct-arco', {login: false, active: 'programmes', active_sub: 'correct-arco', year: new Date().getFullYear(), user: user, notifs: notifs});
    }).catch(err => {
        res.send({
            target: 'database',
            status: false,
            message: 'Unable to connect the database.'
        });
    })
});

module.exports = router;
