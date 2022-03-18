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
const { google } = require("googleapis");
const OAuth2 = google.auth.OAuth2;

const oauth2Client = new OAuth2(
    process.env.APP_CLIENT_ID, // ClientID
    process.env.APP_CLIENT_SECRET, // Client Secret
    "https://developers.google.com/oauthplayground" // Redirect URL
);

oauth2Client.setCredentials({
    refresh_token: process.env.REFRESH_TOKEN
});
const accessToken = oauth2Client.getAccessToken();

//Mailing
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.MAIN_USER,
    pass: process.env.MAIN_PASS,
    clientId: process.env.APP_CLIENT_ID,
    clientSecret: process.env.APP_CLIENT_SECRET,
    refreshToken: process.env.REFRESH_TOKEN,
    accessToken: accessToken
  },
  tls: {
    rejectUnauthorized: false
  }
});


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
    const Limit = 6;
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
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
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
        date_short: date.toLocaleDateString(),
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
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
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
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
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
router.route('/users-list').get(redirectLogin, async (req, res) => {
    const user = req.session.userId;
    mongoose.connect(
        process.env.MONGO_URI,
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
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
 router.route('/manage-access').get(redirectLogin, async (req, res) => {
    const user = req.session.userId;
    mongoose.connect(
        process.env.MONGO_URI,
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
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
 router.route('/edit-user/:email').all(redirectLogin, async (req, res) => {
    const userS = req.session.userId;
    mongoose.connect(
        process.env.MONGO_URI,
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
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
                return res.render('edit-user', {login: false, active: 'users', active_sub: 'users-list', year: new Date().getFullYear(), userEdit: userEdit, user: userEdit, notifs: notifs});
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
router.route('/add-user').post(redirectLogin, (req, res) => {
    mongoose.connect(
        process.env.MONGO_URI,
        {
            useUnifiedTopology: true,
            UseNewUrlParser: true,
        }
    ).then(async () => {
        let user = await {
            username: req.body.username,
            email: req.body.email,
            access: req.body.access,
            password: req.body.password,
        }

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
            res.send({
                status: false,
                icon: 'warning',
                message: 'No file uploaded!'
            }); return;
        } else {
            // use the name of the input field (i.e. "avatar") 
            // to retrieve the uploaded file
            let rh = req.files['rh_file'];
            let salary_up = req.files['salaryup_file'];
            let salary_sheet = req.files['sheet_file'];

            // No RH file selected
            if (!rh && !salary_up) {
                res.send({
                    status: false,
                    icon: 'warning',
                    message: 'No file that contains data uploaded!'
                }); return;
            } 
            // NO Sheet file selected
            if (!salary_sheet) {
                res.send({
                    status: false,
                    icon: 'warning',
                    message: 'No GLOBAL SALARY Sheet file uploaded!'
                }); return;
            }

            let dir = 'uploads'
            // files name
            let rh_filename =  rh ? dir+'/'+rh.name : null;
            let salary_filename =  salary_up ? dir+'/'+salary_up.name : null;
            let sheet_filename = dir+'/'+salary_sheet.name;
            // verifier le repertoire
            if (!fs.existsSync(dir)) {
                await fs.mkdirSync(dir);
            }
            // copier les 3 fichiers
            if (rh) await rh.mv(rh_filename);
            if (salary_up) await salary_up.mv(salary_filename);
            await salary_sheet.mv(sheet_filename);
            // create the output file name
            let date = new Date();
            let opFileName = await `${date.toLocaleDateString().replace(/\//g, '.')} ${salary_sheet.name.split('.xls')[0]} ${date.getTime()}.xlsx`;
            // work books
            var wb_rh = (rh_filename) ? await script.readWBxlsx(rh_filename) : null;
            var wb_salaryup = (salary_filename) ? await script.readWBxlsx(salary_filename) : null;
            var wb_sheet = await script.readWBxlsx(sheet_filename);
            var wb_sheet_style = await script.readWBxlsxstyle(sheet_filename);
            var step = 0;
            if (wb_rh) {
                //.get work sheet rh
                var ws = await script.getWS(wb_rh, 7);
                // check sheets
                if (!ws) {
                    res.send({
                        status: false,
                        icon: 'error',
                        message: 'No specified Sheetname found!'
                    }); return;
                }
                // fetch all data required
                var data = await script.fetchData(ws);
                // if data is empty
                if (data.length <= 0) {
                    res.send({
                        status: false,
                        icon: 'error',
                        message: 'No data found in the RH file! Please verify it.'
                    });
                } else {
                    // output file
                    let output = await script.createOutput(data, wb_sheet, wb_sheet_style);
                    if (output.agent_found === 0) {
                        res.send({
                            status: false,
                            icon: 'error',
                            message: 'No Agent and Required Columns found in the GLOBAL SALARY SHEET! Please verify the file.'
                        });
                    } else {
                        // save file
                        await script.saveFile(output.wb, dir +'/'+ opFileName);
                        step = await step + 1;
                    }
                }
            }
            if (wb_salaryup) {
                let wb = await script.readWBxlsx(salary_filename);

                let ws = await script.getWS(wb, 1);
                let data = await script.getSalaryData(ws);
                if (data.length < 0) {
                    res.send({
                        status: false,
                        icon: 'error',
                        message: 'No data found in the salary file! Please verify the file.'
                    });
                    return;
                } else {
                    // if step one is done change the to the output file.
                    if (step !== 0)
                        wb_sheet = await script.readWBxlsx(dir +'/'+ opFileName);
                    let output = await script.createOutputSalaryUp(data, wb_sheet, wb_sheet_style);
                    
                    await script.saveFile(output.wb, dir +'/'+ opFileName);
                    step = await step + 1;
                }
            }
            if (step > 0) {
                //send response
                await res.send({
                    status: true,
                    icon: 'success',
                    message: 'The file is proccessed successfully.',
                    file: opFileName
                });
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
                        category: 'salary cactulation',
                        description: 'Salary calculation: Recent activity',
                        creation: new Date()
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
                    file: opFileName
                });
            }
            // delete used file after 10 seconds les fichier qu'on viend de telecharger
            await setTimeout(() => {
                fs.readdir(dir, (err, files) => {
                    files.forEach(file => {
                        if (rh && file.includes(rh.name)) 
                            fs.unlinkSync(dir + '/' + file);
                        if (salary_up && file.includes(salary_up.name)) 
                            fs.unlinkSync(dir + '/' + file);
                        if (file.includes(salary_sheet.name)) {
                            fs.unlinkSync(dir + '/' + file);
                        }
                    });
                })
            }, 10000);
            // delete output file after 1 jours
            await setTimeout(() => {
                fs.readdir(dir, (err, files) => {
                    files.forEach(file => {
                        if (file.includes(opFileName)) {
                            fs.unlinkSync(dir + '/' + file);
                        }
                    });
                })
            }, 60*60*24*1000);
        }
    } catch (err) {
        console.log(err)
        res.status(500).send({status: false, icon: 'error', message: 'Server error!'});
    }
});


module.exports = router;
