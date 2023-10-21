require("dotenv").config()
const express = require('express');
const app = express();
const port = 3000;
const xlsx = require('xlsx');
const multer = require('multer');
var fs = require('fs');
const path = require('path');
const mongoose = require('mongoose');
let bodyparser=require('body-parser');
let Express = require('express')

let xltojson = require("C:/Users/praty/OneDrive/Desktop/xlsx/scripts/xltojson.js");

const upload = multer({
  dest: 'uploads/xl/',
  limits: { fileSize: 5 * 1024 * 1024 }, 
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type. Only xlsx files are allowed.'), false);
    }
  },
  storage: multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, 'uploads/xl');
    },
    filename: (req, file, cb) => {
      cb(null, file.originalname); 
    },
  }),
});

// var storage = multer.diskStorage(
//   {
//       destination: './uploads/',
//       filename: function ( req, file, cb ) {
//           cb( null, file.originalname);
//       }
//   }
// );

// var upload = multer( { storage: storage } );
// const upload = multer({ dest: 'uploads/' });

let htmlfolder = path.join(__dirname, "/public/html");


app.use(bodyparser.urlencoded({extended: true}));
app.use(Express.json());
app.use(express.static(__dirname + '/public/css'));
app.set('view engine','ejs');
app.set("views",path.join(__dirname, "./templates/views"))

const nodesSchema = {
  title: String,
  content: String,
  type: String
}

const Node = mongoose.model('excel', nodesSchema)

app.get("/",(req,res) => {
  const directoryPath = path.join(__dirname, '/uploads/xl');
  let fileList = [];
  fs.readdir(directoryPath, function (err, files) {
      if (err) {
          return console.log('Unable to scan directory: ' + err);
      }
      files.forEach(function (file) {
        fileList.push(file.substring(0,file.length-5));
      });
      res.render('index.ejs',{
        list : fileList,
      });
  });
});

app.post('/upload', upload.single('file'), (req, res) => {
  const file = req.file;
  if (!file) {
    return res.redirect("/");
  }
  try {
    xltojson.xtj(file);
    res.redirect("/");
  } 
  catch (error) {
    console.error('Error reading file:', error);
    res.status(500).send('Error reading file.');
  }
});

app.post('/content',(req,res) => {
  filePath = path.join(__dirname, '/uploads/json/'+req.body.fileName+".json");
  fs.readFile(filePath, 'utf8', (err, data) => {
      if (err) {
        console.error(err);
      } 
      else {
        try {
          res.render('content.ejs',{
              data : data,
          });
        } catch (error) {
          console.error('Error parsing JSON:', error);
        }
      }
    });
});

app.post('/delete',(req,res) => {
  jsonFilePath = path.join(__dirname, '/uploads/json/'+req.body.fileName+".json");
  xlFilePath = path.join(__dirname, '/uploads/xl/'+req.body.fileName+".xlsx");
  fs.unlinkSync(jsonFilePath);
  fs.unlinkSync(xlFilePath);
  res.redirect("/");
});

app.post('/download',(req,res) => {
  xlFilePath = path.join(__dirname, '/uploads/xl/'+req.body.fileName+".xlsx");
  res.status(200).download(xlFilePath, req.body.fileName+".xlsx");
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
  mongoose.connect(process.env.dburi)
  .then((result)=>console.log("mdb 1 passed"))
  .catch((err)=>console.log(err));
});