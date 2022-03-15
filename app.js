require('dotenv').config();
const express = require('express');
const expressLayouts = require('express-ejs-layouts');
const expressUpload = require('express-fileupload');
const router = require('./controllers/controller');
const app = express();
const bodyParser = require('body-parser');
const PORT = process.env.PORT || 8086;
const session = require('express-session');

// static files 
app.use(express.static('public'));
app.use(express.static('uploads'));
app.use(express.static('template'));
app.use('/css', express.static(__dirname + 'public/css'));
app.use('/js', express.static(__dirname + 'public/js'));
app.use(express.static(__dirname))
app.use(session({
  name: 'sid',
  resave: true,
  saveUninitialized: true,
  secret: 'psc',
}));

// use upload
app.use(expressUpload());

// set template enginea
app.use(expressLayouts)
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/', router);

const server = app.listen(PORT, () => {
  const port = server.address().port;
  console.log(`Express is working on port ${port}`);
});