import { stringify } from 'querystring';
import { Workbook, Worksheet, Row, ValueType, CellValue } from 'exceljs';
import { Entry, Hours } from './entry';

export class DSW {
    filepath:string
    AllEntries:Entry[]
    headerMarker:number
    startMarker:number
    endMarker:number
    columnList:string[]
    workSheet:Worksheet
    columnNumbers:{[key: string]:number}

    constructor(fp:string) {
        this.filepath = fp;
        this.AllEntries = [];
        this.startMarker = 0;
        this.endMarker = 0;
        this.columnList = [];
        this.columnNumbers = {};
    }

    async init() {
        const AllEntries = [];

        const workbook = new Workbook();
        await workbook.xlsx.readFile(this.filepath);

        const dsw = workbook.getWorksheet(1);

        this.workSheet = dsw;

        [this.headerMarker, this.endMarker] = getMarkers(dsw);
        this.startMarker = this.headerMarker + 1;

        this.scanColumns(dsw);

        this.parseData(dsw);

        // console.log("Column List:", this.columnList);
        // console.log("All Entries:", this.AllEntries);
    }

    scanColumns = function(dsw:Worksheet) {
        for(let i = 1; i < dsw.actualColumnCount; i++) {
            const value = dsw.getRow(this.headerMarker).getCell(i).value;
            let name = "";
            switch(value) {
                case "Date": name = "Date";
                             break;

                case "WA Name": name = "WAName";
                             break;

                case "Svc Area #": name = "SvcAreaNumber";
                             break;

                case "Veh #": name = "VehicleNumber";
                             break;

                case "Driver/ Helper Name": name = "DriverHelperName";
                             break;

                case "WA#": name = "WASNumber";
                             break;

                case "MDSG": name = "MDSG";
                             break;

                case "DST": name = "DST";
                             break;

                case "VScan Pkgs": name = "VScanPkgs";
                             break;

                case "Del Stps": name = "DelStps";
                             break;

                case "PU Stps": name = "PUStps";
                             break;

                case "DIFF": name = "DIFF";
                             break;

                case "Act Del Stps": name = "ActDelStps";
                             break;

                case "Act Del Pkgs": name = "ActDelPkgs";
                             break;

                case "Act PU Stps": name = "ActPUStps";
                             break;

                case "Act PU Pkgs": name = "ActPUPkgs";
                             break;

                case "ILS%": name = "ILSPercent";
                             break;

                case "ILS Impact Pkgs": name = "ILSImpactPkgs";
                             break;

                case "Non delvd Stps": name = "NondelvdStps";
                             break;

                case "Code 85": name = "Code85";
                             break;

                case "All Status Code Pkgs": name = "AllStatusCodePkgs";
                             break;

                case "P'L M'L": name = "PLML";
                             break;

                case "DNA": name = "DNA";
                             break;

                case "Snd Agn": name = "SndAgn";
                             break;

                case "Exc's": name = "Excs";
                             break;

                case "VSA vs STAR (DIFF)": name = "VSAvsSTARDIFF";
                             break;

                case "% Returns Scans": name = "ReturnsScansPercent";
                             break;

                case "Miles": name = "Miles";
                             break;

                case "On Road Hours": name = "OnRoadHours";
                             break;

                case "On Duty Hours": name = "OnDutyHours";
                             break;

                case "Pot. DOT Hrs Viols": name = "PotDOTHrsViols";
                             break;

                case "Next Avail On Duty": name = "NextAvailOnDuty";
                             break;

                case "Pot. Miss PUs": name = "PotMissPUs";
                             break;

                case "E/L PUs": name = "ELPUs";
                             break;

                case "Req. Sig.": name = "ReqSig";
                             break;

                case "Date Certain": name = "DateCertain";
                             break;

                case "Evening": name = "Evening";
                             break;

                case "Appt": name = "Appt";
                             break;
            }
            console.log('Column Name: ', name)
            this.columnNumbers[name] = i;
            this.columnList[i-1] = name;
        }

        console.log('Column Numbers: ', this.columnNumbers);
    }

    parseData = function(dsw:Worksheet) {
        let savedStateEntry:Entry;
        let debug = false;
        for(let i = this.startMarker;i < this.endMarker; i++) {
            // console.log(i);
            const row = dsw.getRow(i);
            const nextRow = dsw.getRow(i+1);
            const date = row.getCell(1).value;

            const DriverHelperName = this.parseStringValue(row,5);

            if(DriverHelperName === '') {
                debug = false;
            } else {
                debug = false;
            }

            const nextRowDate = nextRow.getCell(1).value;
            if(nextRowDate === '' && date !== '') {
                savedStateEntry = this.parseEntry(row);
                // console.log(i, 'Next row date blank, skipping')
                if(debug) console.log(savedStateEntry);
                continue;
            }

            const entry = this.parseEntry(row, savedStateEntry);
            if(debug) console.log(entry);
            this.AllEntries.push(entry);
            // console.log(date, entry);
        }
    }

    parseEntry = (row, savedStateEntry):Entry =>{
        let entry = new Entry();

        const date = row.getCell(1).value;

        if(date === '') {
            // entry = JSON.parse(JSON.stringify(savedStateEntry));
            entry = {...savedStateEntry};
        }

        if(date !== '') {
            // This is a helper
            this.parseDate(entry, row);
            entry.WAName = this.parseStringValue(row, this.columnNumbers.WAName);
            entry.MDSG = this.parseStringValue(row,this.columnNumbers.MDSG);
            entry.DST = this.parseIntValue(row, this.columnNumbers.DST);
            entry.VScanPkgs = this.parseIntValue(row, this.columnNumbers.VScanPkgs);
            entry.DelStps = this.parseIntValue(row, this.columnNumbers.DelStps);
            entry.PUStps = this.parseIntValue(row, this.columnNumbers.PUStps);
            entry.DIFF = this.parseIntValue(row, this.columnNumbers.DIFF);
            entry.ILSPercent = this.parseILSPercent(entry, row);
            entry.ILSImpactPkgs = this.parseIntValue(row, this.columnNumbers.ILSImpactPkgs);
            entry.NondelvdStps = this.parseIntValue(row, this.columnNumbers.NondelvdStps);
            entry.Code85 = this.parseIntValue(row, this.columnNumbers.Code85);
            entry.AllStatusCodePkgs = this.parseIntValue(row, this.columnNumbers.AllStatusCodePkgs);
            entry.PLML = this.parseIntValue(row, this.columnNumbers.PLML);
            entry.DNA = this.parseIntValue(row, this.columnNumbers.DNA);
            entry.SndAgn = this.parseIntValue(row, this.columnNumbers.SndAgn);
            entry.Excs = this.parseIntValue(row, this.columnNumbers.Excs);
            entry.VSAvsSTARDIFF = this.parseIntValue(row, this.columnNumbers.VSAvsSTARDIFF);
            // entry.ReturnsScansPercent = this.parseReturnsScansPercent(entry, row);
            entry.Miles = this.parseIntValue(row, this.columnNumbers.Miles);
            entry.PotDOTHrsViols = this.parseStringValue(row, this.columnNumbers.PotDOTHrsViols);
            entry.NextAvailOnDuty = this.parseHours(row.getCell(this.columnNumbers.NextAvailOnDuty).value);
            entry.PotMissPUs = this.parseIntValue(row, this.columnNumbers.PotMissPUs);
            entry.ELPUs = this.parseIntValue(row, this.columnNumbers.ELPUs);
            entry.ReqSig = this.parseIntValue(row, this.columnNumbers.ReqSig);
            entry.DateCertain = this.parseIntValue(row, this.columnNumbers.DateCertain);
            entry.Evening = this.parseIntValue(row, this.columnNumbers.Evening);
            entry.Appt = this.parseIntValue(row, this.columnNumbers.Appt);
        }

        let celloffset:number = 0;
        if (date === '' &&
            (row.getCell(this.columnNumbers.ActDelStps).value == null || row.getCell(this.columnNumbers.ActDelStps).value === '')
            ) {
                // This is helper and cells have been moved to right by 1
                celloffset = 1;
            }
        entry.SvcAreaNumber = this.parseIntValue(row, this.columnNumbers.SvcAreaNumber);
        entry.VehicleNumber = this.parseVehicleNumber(row.getCell(this.columnNumbers.VehicleNumber).value);
        entry.DriverHelperName = this.parseStringValue(row,this.columnNumbers.DriverHelperName);
        entry.WANumber = this.parseIntValue(row, this.columnNumbers.WANumber);
        entry.ActDelStps = this.parseIntValue(row, this.columnNumbers.ActDelStps, celloffset);
        entry.ActDelPkgs = this.parseIntValue(row, this.columnNumbers.ActDelPkgs, celloffset);
        entry.ActPUStps = this.parseIntValue(row, this.columnNumbers.ActPUStps, celloffset);
        entry.ActPUPkgs = this.parseIntValue(row, this.columnNumbers.ActPUPkgs, celloffset);
        entry.OnRoadHours = this.parseHoursFromRow(row, this.columnNumbers.OnRoadHours, celloffset);
        // parseInt(row.getCell(29).value.toString()) | 0;
        entry.OnDutyHours = this.parseHoursFromRow(row, this.columnNumbers.OnDutyHours, celloffset);
        // parseInt(row.getCell(30).value.toString()) | 0;

        return entry;
    }

    parseDate = (entry:Entry, row:Row) => {
        // console.log('parsing date')
        const datevalue = row.getCell(this.columnNumbers.Date).value;
        entry.Date = new Date();
        const year = parseInt(datevalue.toString().substr(0,4), 10);
        const month = parseInt(datevalue.toString().substr(4,2), 10);
        const date = parseInt(datevalue.toString().substr(6,2), 10);
        entry.Date.setFullYear(year, month-1,date);
        entry.Date.setHours(0);
        // console.log('end parsing date')
    }

    parseILSPercent = (entry:Entry, row:Row):number => {
        let ils = row.getCell(this.columnNumbers.ILSPercent).value.toString();
        ils = ils.split("%").join("");
        return parseInt(ils, 10) || 0;
    }

    parseHoursFromRow = (row:Row, cellNumber:number, cellOffset:number):Hours  => {
        if (!cellNumber) {
            return null;
        }
        if (!cellOffset) {
            cellOffset = 0;
        }
        const localCellNumber = cellNumber + cellOffset;
        const hours = {hours: 0, minutes: 0};
        const value = row.getCell(localCellNumber).value;
        if(value != null) {
            return this.parseHours(value);
        }
        return hours;
    }

    parseHours = (value:CellValue):Hours => {
        const hours = {hours: 0, minutes: 0}
        if(value == null) {
            return hours;
        }

        const valueString = value.toString()
        if(value.toString() === "") {
            return hours;
        }

        hours.hours = parseInt(valueString.split(":")[0], 10);
        hours.minutes = parseInt(valueString.split(":")[1], 10);

        return hours
    }

    parseStringValue = (row:Row,cellNumber:number):string => {
        if (!cellNumber) {
            return "";
        }
        if(row.getCell(cellNumber).value != null) {
            return row.getCell(cellNumber).value.toString()
        } else {
            return "";
        }

    }

    parseIntValue = (row:Row, cellNumber:number, cellOffset?:number):number => {
        if (!cellNumber) {
            return 0;
        }
        if (!cellOffset) {
            cellOffset = 0;
        }
        const localCellNumber = cellNumber + cellOffset;
        if(row.getCell(localCellNumber).value == null) {
            return 0;
        } else if(row.getCell(localCellNumber).value.toString() === "") {
            return 0;
        } else {
            return parseInt(row.getCell(localCellNumber).value.toString(), 10);
        }
    }

    parseVehicleNumber = (value:ValueType):number => {
        if(value == null) {
            return 0
        }
        const valueString = value.toString();
        const valueSplit = valueString.split(" ");
        if(valueSplit.length > 0) {
            return parseInt(valueSplit[valueSplit.length - 1], 10);
        }

    }
};


const getMarkers = (dsw):[number, number] => {
    let startMarker = 0;
    let endMarker = 0;

    for (let i = 0; i < dsw.rowCount; i++) {
        const row = dsw.getRow(i);
        if (row.getCell(1).value != null) {
            if (row.getCell(1).value.indexOf("Date") !== -1) {
                startMarker = i;
                break;
            }
        }
    }
    for (let i = 0; i < dsw.rowCount; i++) {
        const row = dsw.getRow(i);
        if (row.getCell(1).value != null) {
            if (row.getCell(1).value.indexOf("Contract") !== -1 && row.getCell(1).value.indexOf("Total") !== -1) {
                endMarker = i;
                break;
            }
        }
    }

    if(startMarker >= dsw.rowCount) {
        throw new Error("Cannot find start marker");
    }
    if(endMarker >= dsw.rowCount) {
        throw new Error("Cannot find end marker");
    }

    console.log("startMarker:", startMarker);
    console.log("endMarker:", endMarker);
    return [startMarker, endMarker];
};