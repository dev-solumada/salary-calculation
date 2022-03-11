require('dotenv').config();
const express = require('express');
const router = express.Router();
const fs = require('fs');
const script = require('./script.js')
const mongoose = require('mongoose');
const UserSchema = require('../models/UserSchema');
const SCSchema = require('../models/SCSchema');
const nodemailer = require('nodemailer');
const moment = require('moment');

//Mailing
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.MAIN_USER,
    pass: process.env.MAIN_PASS,
  },
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
router.route('/logout').all(redirectLogin, (req, res) => {
    req.session.destroy();
    return res.redirect('/login');
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
        csInfoActivity: csInfoActivity
    });
});

/**
 * PROGRAMME - Salary Calculation.
 */
router.route('/salary-calculation').get(redirectLogin, (req, res) => {
    const user = req.session.userId;
    res.render('salary-calculation', {login: false, active: 'programmes', active_sub: 'salary-calculation', year: new Date().getFullYear(), user: user});
});

/**
 * USERS - Add new user.
 */
router.route('/add-new-user').get(redirectLogin, (req, res) => {
    const user = req.session.userId;
    res.render('add-new-user', {login: false, active: 'users', active_sub: 'add-new-user', year: new Date().getFullYear(), user: user});
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
        return res.render('users-list', {login: false, active: 'users', active_sub: 'users-list', year: new Date().getFullYear(), users: users, user: user});
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
        res.render('manage-access', {login: false, active: 'users', active_sub: 'manage-access', year: new Date().getFullYear(), users: users, user: user});
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
                await res.send({
                    status: true,
                    message: 'User successfully updated!'
                });
            }
        } else { // GET
            // find all users
            let userEdit = await UserSchema.findOne({email: req.params.email});
            if (userEdit) {
                return res.render('edit-user', {login: false, active: 'users', active_sub: 'users-list', year: new Date().getFullYear(), userEdit: userEdit, user: userEdit});
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
                    if (error) {
                        console.log(error);
                    }
                });
            }
            
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
            let salary_sheet = req.files['sheet_file'];

            // No RH file selected
            if (!rh) {
                res.send({
                    status: false,
                    icon: 'warning',
                    message: 'No RH file uploaded!'
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
            let rh_filename = dir+'/'+rh.name;
            let sheet_filename = dir+'/'+salary_sheet.name;
            // verifier le repertoire
            if (!fs.existsSync(dir)) {
                await fs.mkdirSync(dir);
            }
            // copier les 2 fichiers
            await rh.mv(rh_filename);
            await salary_sheet.mv(sheet_filename);
            // create the output file name
            let date = new Date();
            let opFileName = await `${salary_sheet.name.split('.xls')[0]}_${String(date.getDay()).padStart(2, '0')}${String(date.getMonth()).padStart(2, '0')}${date.getFullYear()}.xlsx`;
            // work books
            var wb_rh = await script.readWBxlsx(rh_filename);
            var wb_sheet = await script.readWBxlsx(sheet_filename);
            var wb_sheet_style = await script.readWBxlsxstyle(sheet_filename);
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
                        
                    }).catch(err => {
                        res.send({
                            target: 'database',
                            status: false,
                            message: 'Unable to connect the database.'
                        });
                    })
                }
            }
            // delete used file after 30 seconds
            await setTimeout(() => {
                fs.readdir(dir, (err, files) => {
                    files.forEach(file => {
                        if (file.includes(rh.name) || file.includes(salary_sheet.name) || file.includes(opFileName)) {
                            fs.unlinkSync(dir + '/' + file);
                        }
                    });
                })
            }, 5000 * 6);
        }
    } catch (err) {
        console.log(err)
        res.status(500).send({status: false, icon: 'error', message: 'Server error!'});
    }
});


module.exports = router;