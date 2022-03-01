const express = require('express')
const app = express()
const bodyparser = require('body-parser')
const fs = require('fs');
const csv = require('fast-csv');
const readXlsxFile = require('read-excel-file/node');
const mysql = require('mysql')
const multer = require('multer')
const path = require('path')

//use express static folder or insted of this we create public folder in our node app
app.use(express.static("./public"))
// body-parser middleware use
app.use(bodyparser.json())
app.use(bodyparser.urlencoded({
    extended: true
}))
// Database connection
const db = mysql.createConnection({
    connectionLimit: 10, //datbase pooling
    host: "localhost",
    user: "root",
    password: "",
    database: "csvtomysql"
})
db.connect(function (err) {
    if (err) {
        return console.error('error: ' + err.message + err.stack);
    }
    console.log('Connected to the MySQL server.Connected Id:- ' + db.threadId);
})
//! Use of Multer
var storage = multer.diskStorage({
    destination: (req, file, callBack) => {
        if (file.fieldname === "uploadfile") {
            callBack(null, './uploads/')
        }
    },
    filename: (req, file, callBack) => {
        callBack(null, file.fieldname + '-' + Date.now() + path.extname(file.originalname))
    }
})
var upload = multer({
    storage: storage,
 });
//@type   POST
//route for post data
app.post('/uploadfile', upload.single("uploadfile"), (req, res) => {
    UploadCsvDataToMySQL(__dirname + '/uploads/' + req.file.filename);
    res.json({
        'msg': 'csv File uploaded/import successfully!', 'file': req.file
    });
});
function UploadCsvDataToMySQL(filePath) { 
    let stream = fs.createReadStream(filePath);
    let csvData = [];
    let csvStream = csv
        .parse()
        .on("data", function (data) {
            csvData.push(data);
        })
        .on("end", function () {
            // Remove Header ROW
            csvData.shift();

            let query = 'INSERT INTO user (Id, Name, Email, Phonenumber) VALUES ?';
            db.query(query, [csvData], (error, response) => {
                console.log(error || response);
            });


            // delete file after saving to MySQL database
            //fs.unlinkSync(filePath)
        });
    stream.pipe(csvStream);
}

//excel to myql

var store = multer.diskStorage({
    destination: (req, file, cb) => {
        if (file.fieldname === "renfile") {
            cb(null, './upload/')
        }
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + '-' + Date.now() + path.extname(file.originalname))
    }

})
var upload = multer({
     storage:store
});
//! Routes start
//route for Home page
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

// -> Express Upload RestAPIs
app.post('/upload', upload.single('renfile'), (req, res) => {
    importExcelData2MySQL(__dirname + '/upload/' + req.file.filename);
    res.json({
        'msg': 'excel File uploaded/import successfully!', 'file': req.file
    });
});
// -> Import Excel Data to MySQL database
function importExcelData2MySQL(filePath) {
    // File path.
    readXlsxFile(filePath).then((rows) => {
        // `rows` is an array of rows
        // each row being an array of cells.     
        console.log(rows);
        /**
        [ [ 'Id', 'Name', 'Address', 'Age' ],
        [ 1, 'john Smith', 'London', 25 ],
        [ 2, 'Ahman Johnson', 'New York', 26 ]
        */
        // Remove Header ROW
        rows.shift();
        // Open the MySQL connection
        let queryii = 'INSERT INTO tutorials (id,title,description) VALUES ?';
        db.query(queryii, [rows], (error, response) => {
            console.log(error || response);
            /**
            OkPacket {
            fieldCount: 0,
            affectedRows: 5,
            insertId: 0,
            serverStatus: 2,
            warningCount: 0,
            message: '&Records: 5  Duplicates: 0  Warnings: 0',
            protocol41: true,
            changedRows: 0 }  
            */
        });
    })
}

//create connection
const PORT = process.env.PORT || 5001
app.listen(PORT, () => console.log(`Server is running at port ${PORT}`))