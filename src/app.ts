import { DSW } from './dsw';
import { Report } from './report';
import * as fs from 'fs';
import * as express from 'express';
import * as path from 'path';
import * as bodyParser from 'body-parser';
import { Employee } from './employee';

const app = express();

let dswFileNames = [];
let reportFileNames = [];
let employeeFileNames = [];
let employees = new Employee();
const dswFilesPath = path.join(path.dirname(process.execPath), '..', 'dsw');
const reportFilesPath = path.join(path.dirname(process.execPath), '..', 'report');
const employeeFilesPath = path.join(path.dirname(process.execPath), '..', 'employees');

fs.mkdirSync(dswFilesPath, { recursive: true })
fs.mkdirSync(reportFilesPath, { recursive: true })
fs.mkdirSync(employeeFilesPath, { recursive: true })

async function generateEmployeeReport(inputFilename, outputFilename) {
    const d = new DSW(inputFilename);
    await d.init()

    const r = new Report(d, employees);
    const employeeReport = r.employeeReport()
    r.writeEmployeeReport(employeeReport, outputFilename)
    console.log('Finished writing employee report');
    // console.log(JSON.stringify(employees,null, 2))

}

function setupRoutes() {
    // app.use(bodyParser.json());
    app.use(bodyParser.json({limit: '50mb'}));
    app.use(bodyParser.urlencoded({limit: '50mb', extended: true}));

    app.get("/filenames", (req, res) => {
        // console.log("dswFileNames ", dswFileNames);
        // console.log("reportFileNames ", reportFileNames);
        // console.log("employeeFileNames ", employeeFileNames);
        res.send({ dswFileNames, reportFileNames, employeeFileNames });
    });

    app.get('/employees', (req, res) => {
        res.send({ employees: employees.List })
    });

    // login route --setup database for this
    app.post("/login", (req, res) => {
        console.log("request ", req.body);
        const { email, password, captcha } = req.body.newUser;
        // let recaptcha_url = `https://www.google.com/recaptcha/api/siteverify?secret=${cmdArgs.SECRETKEY}&response=${captcha}&remoteip=${req.connection.remoteAddress}`;
        // request(recaptcha_url, (error, resp, body) => {
        //     body = JSON.parse(body);
        //     if (body.success !== undefined && !body.success) {
        //         res.send(new Error("Captcha validation failed. If JavaScript is disabled in your browser, then please enable it and try again."));
        //     } else {
        if (email === "abc@xyz.com" && password === "@bc1234") {
            console.log('Login successful');
            res.send({
                firstname: "Abc",
                expiresIn: "7200"
            });
        } else {
            console.log('Invlid username/password');
            res.send(new Error("Username or Password is incorrect"));
        }
        // }
        // });
    });

    app.post("/employeeFileUpload", (req,res) => {
        console.log("req ", req.body);
        const fileName = req.body.fileName;
        const fileContents = req.body.fileContents;
        res.send("Received File");
    });

    app.post("/dswFileUpload", (req,res) => {
        console.log("req ", req.body);
        const fileName = req.body.fileName;
        const fileContents = req.body.fileContents;
        res.send("Received File");
    });

    app.post("/payrollFileUpload", (req,res) => {
        console.log("req ", req.body);
        const fileName = req.body.fileName;
        const fileContents = req.body.fileContents;
        res.send("Received File");
    });

    app.post("/download", (req,res) => {
        console.log("req.body ", req.body)
        const fileName = req.body.data.file;
        const folder = req.body.data.folder.slice(0,req.body.data.folder.indexOf("Files"));
        console.log(fileName,folder);
        if(folder === "employee"){
            const downloadPath = employeeFilesPath + "/" + fileName;
            const readStream = fs.createReadStream(downloadPath);
            res.writeHead(200, {'Content-disposition': 'attachment; filename=employee-list.xlsx', 	'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'responseType': 'arraybuffer' });
            readStream.pipe(res)
            // res.download(downloadPath);
        } else if(folder === "dsw"){
            const downloadPath = dswFilesPath + "/" + fileName;
            res.download(downloadPath);
        }
    })

    app.post("/delete", (req,res) => {
        console.log("req.body ", req.body)
        const fileName = req.body.data.file;
        const folder = req.body.data.folder.slice(0,req.body.data.folder.indexOf("Files"));
        console.log(fileName,folder);
        if(folder === "employee"){
            const deletePath = employeeFilesPath + "/" + fileName;
            fs.unlinkSync(deletePath);
            res.send("Deleted Successfully");
        } else if(folder === "dsw"){
            const deletePath = dswFilesPath + "/" + fileName;
            fs.unlinkSync(deletePath);
            res.send("File Deleted  Successfully");
        } else if(folder === "report"){
            const deletePath = reportFilesPath + "/" + fileName;
            fs.unlinkSync(deletePath);
            res.send("File Deleted  Successfully");
        } else {
            res.send(new Error("File not found"));
        }
    })

    if (process.argv.length === 2) {
        app.listen(5000, () => {
            console.log("Server is running on port " + 5000);
        });
    }
}

function readInputFiles() {
    fs.readdir(dswFilesPath, (err, files) => {
        const usefulFiles = files.filter((f) => {
            return f !== '.gitkeep' && !f.startsWith("~");
        })
        dswFileNames = usefulFiles;
    });
}

function readReportFiles() {
    fs.readdir(reportFilesPath, (err, files) => {
        const usefulFiles = files.filter((f) => {
            return f !== '.gitkeep' && !f.startsWith("~");
        })
        reportFileNames = usefulFiles;
    });
}

function readEmployeeFiles() {
    fs.readdir(employeeFilesPath, (err, files) => {
        const usefulFiles = files.filter((f) => {
            return f !== '.gitkeep' && !f.startsWith("~");
        })
        employeeFileNames = usefulFiles;
    });
}

async function start() {

    console.log("argv", process.argv)
    let paramsOffset = 0;
    if(process.argv[0].endsWith(".exe")) {
        paramsOffset = 0;
    }

    if (process.argv.length === 4) {
        if(!process.argv[2-paramsOffset].endsWith(".xlsx")) {
            console.log("The provided employee file should be 'xlsx' type")
            process.exit(1)
        }

        employees = new Employee(`${process.argv[2-paramsOffset]}`);
        await employees.init();
        let inputFilename = process.argv[3-paramsOffset]
        if(!path.isAbsolute(inputFilename)) {
            inputFilename = path.join(path.dirname(process.execPath), inputFilename)
        }
        console.log("Provided input file path:", inputFilename)
        if(!inputFilename.endsWith(".xlsx")) {
            console.log("The provided input file should be 'xlsx' type")
            process.exit(1)
        }
        // const inputFilename = process.argv[3] + '.xlsx';
        const filenamewithnoext = inputFilename.replace(".xlsx", "")
        // const outputFilename = process.argv[3] + '-report.xlsx';
        const outputFilename = filenamewithnoext + '-report.xlsx';
        await generateEmployeeReport(inputFilename, outputFilename);
        return;
    } else {
        console.log("ERROR: Invalid arguments. Please provide the WSW path ")
    }

    employees = new Employee('./employees/employee-list.xlsx');
    await employees.init();

    if (process.argv.length === 3) {
        // const inputFilename = process.argv[2] + '.xlsx';
        // const outputFilename = process.argv[2] + '-report.xlsx';
        let inputFilename = process.argv[2-paramsOffset]
        if(!path.isAbsolute(inputFilename)) {
            inputFilename = path.join(path.dirname(process.execPath), inputFilename)
        }
        console.log("Provided input file path:", inputFilename)
        if(!inputFilename.endsWith(".xlsx")) {
            console.log("The provided input file should be 'xlsx' type")
            process.exit(1)
        }
        // const inputFilename = process.argv[3] + '.xlsx';
        const filenamewithnoext = inputFilename.replace(".xlsx", "")
        // const outputFilename = process.argv[3] + '-report.xlsx';
        const outputFilename = filenamewithnoext + '-report.xlsx';

        await generateEmployeeReport(inputFilename, outputFilename);
        return;
    }

    // Exit if above conditions not met
    console.log("No arguments provided. Please provide the WSW file location")
    return;

    readInputFiles();
    readReportFiles();
    readEmployeeFiles();

    fs.watch('./dsw/', (eventType, filename) => {
        readInputFiles();

        if (eventType !== 'change' || filename === '.gitkeep' || filename.startsWith("~") ) {
            return;
        }

        console.log("Event type: ", eventType);

        const inputFile = path.join(dswFilesPath, filename)
        try {
            fs.statSync(inputFile);
        } catch (e) {
            if (e.code === 'ENOENT') {
                // File deleted
            } else {
                console.log("ERROR:", e.code);
            }
            return;
        }

        if (filename !== '.' && filename !== '..') {
            const outputFilename = path.join(reportFilesPath, filename);

            // console.log("Input:", inputFile, "output: ", outputFilename);
            generateEmployeeReport(inputFile, outputFilename);
        }
    });

    fs.watch(reportFilesPath, (eventType, filename) => {
        readReportFiles();
    });

    fs.watch(employeeFilesPath, (eventType, filename) => {
        readEmployeeFiles();
    });
}

// setupRoutes();
start();
