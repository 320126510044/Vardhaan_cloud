var express = require("express");
var multer = require("multer");
const ExcelJS = require("exceljs/dist/es5");
const fs = require("fs");
const path = require("path");
var port = 3000;

var app = express();

let workbook = new ExcelJS.Workbook();
let worksheet = workbook.addWorksheet("pdfListSheet");
worksheet.columns = [{ header: "filename", key: "id" }];

var storage = multer.diskStorage({
  filename: function (req, file, cb) {
    cb(null, file.originalname);
    worksheet.addRow({ id: file.originalname });
  },
});
var upload = multer({ storage: storage });

/*
app.use('/a',express.static('/b'));
Above line would serve all files/folders inside of the 'b' directory
And make them accessible through http://localhost:3000/a.
*/
app.use(express.static(__dirname + "/public"));
app.use("/uploads", express.static("uploads"));

app.post(
  "/profile-upload-single",
  upload.single("profile-file"),
  function (req, res, next) {
    // req.file is the `profile-file` file
    // req.body will hold the text fields, if there were any
    console.log(JSON.stringify(req.file));
    var response = '<a href="/">Home</a><br>';
    response += "Files uploaded successfully.<br>";
    response += `<img src="${req.file.path}" /><br>`;

    return res.send(response);
  }
);

app.post(
  "/profile-upload-multiple",
  upload.array("profile-files", 12),
  async function (req, res, next) {
    // req.files is array of `profile-files` files
    // req.body will contain the text fields, if there were any
    console.log(JSON.stringify(req.file));

    await workbook.xlsx.writeFile("uploads/hello.xlsx");
    return res.status(200);
  }
);

app.listen(process.env.PORT || 3000 , () =>
  console.log(`Server running on port ${port}!\nClick http://localhost:3000/`)
);

//  const workbook = new ExcelJS.Workbook();
//  const worksheet = workbook.addWorksheet("pdfListSheet");
//  worksheet.columns = [
//    { header: "Id", key: "id" },
//    { header: "Name", key: "name" },
//    { header: "Age", key: "age" },
//  ];
//  const row = worksheet.addRow({ id: 1, name: "John Doe", age: 35 });

//  await workbook.xlsx.writeFile("hello.xlsx");

app.post("/delete", async (req, res) => {
  const dirPath = "./uploads";
  fs.unlink("./uploads/hello.xlsx", (err) => {
    if (err) {
      console.log(err);
    }
  });
  return res.status(200);
});
