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
const extendTimeoutMiddleware = (req, res, next) => {
  const space = ' ';
  let isFinished = false;
  let isDataSent = false;

  // Only extend the timeout for API requests
  if (!req.url.includes('/api')) {
    next();
    return;
  }

  res.once('finish', () => {
    isFinished = true;
  });

  res.once('end', () => {
    isFinished = true;
  });

  res.once('close', () => {
    isFinished = true;
  });

  res.on('data', (data) => {
    // Look for something other than our blank space to indicate that real
    // data is now being sent back to the client.
    if (data !== space) {
      isDataSent = true;
    }
  });

  const waitAndSend = () => {
    setTimeout(() => {
      // If the response hasn't finished and hasn't sent any data back....
      if (!isFinished && !isDataSent) {
        // Need to write the status code/headers if they haven't been sent yet.
        if (!res.headersSent) {
          res.writeHead(202);
        }

        res.write(space);

        // Wait another 15 seconds
        waitAndSend();
      }
    }, 15000);
  };

  waitAndSend();
  next();
};

app.use(extendTimeoutMiddleware);

// use upload
app.use(expressUpload());

// set template enginea
app.use(expressLayouts)
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
const server = require("http").createServer(app);
const io = require("socket.io")(server);
io.on('connection', socket => {
  socket.on('server', mess => {
    socket.emit('server', mess);
    console.log(mess)
  });
  app.set("socket", socket);
});
app.set("io", io);

app.use('/', router);

server.listen(PORT, () => {
  const port = server.address().port;
  console.log(`Express is working on port ${port}`);
});
