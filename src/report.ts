import { DSW } from "./dsw";
import { Hours } from "./entry";
import { Workbook, Worksheet, Column, CellFormulaValue } from "exceljs";
import { EmployeeSummary } from "./employeesummary";
import * as fs from 'fs';
import { Employee } from "./employee";

export class Report {
    dsw: DSW;
    employeeList: Employee;
    constructor(dsw: DSW, employeelist: Employee) {
        this.dsw = dsw;
        this.employeeList = employeelist;
    }

    calculateHourlyPay = (hourlyRate: number, hours: number | Hours): number => {
        let pay = 0;
        hourlyRate = hourlyRate || 11;

        let h = 0;
        if (typeof (hours) === 'object') {
            h = hours.hours + hours.minutes / 60;
        } else if (typeof (hours) === 'number') {
            h = hours
        } else {
            console.log("hours is", typeof (hours))
            return 0
        }

        if (hours >= 40) {
            pay += (40 * hourlyRate)
            pay += (h - 40) * (hourlyRate * 1.5)
        } else {
            pay += (h * hourlyRate)
        }

        return pay
    }

    calculateStopPay = (stopRate: number, stops: number): number => {
        stopRate = stopRate || 1.30;
        stops = stops || 0;
        // console.log("Stop Rate:", stopRate, "stops", stops)
        return stops * stopRate
    }

    employeeReport = (): object => {
        // console.log("DSW:", this.dsw);
        const allEntries = this.dsw.AllEntries;

        const employees: Map<string, EmployeeSummary> = new Map<string, EmployeeSummary>();

        // console.log(allEntries);
        for (const entry of allEntries) {
            const employeename = entry.DriverHelperName.toLowerCase();
            if (employees[employeename] === undefined) {
                employees[employeename] = new EmployeeSummary();
            };
            employees[employeename].Name = entry.DriverHelperName;
            employees[employeename].addEntry(entry);
        }

        for (const employeename in employees) {
            if (employees.hasOwnProperty(employeename)) {
                employees[employeename].HourlyPay = this.calculateHourlyPay(employees[employeename].HourlyRate, employees[employeename].Totals.OnDutyHours)
                employees[employeename].StopPay = this.calculateStopPay(employees[employeename].StopRate, employees[employeename].Totals.AllStops)
            }
        }

        // console.log(employees[''])
        return employees
    }

    writeEmployeeReport = (employees: object, filepath: string) => {
        fs.writeFileSync("temp.json", JSON.stringify(employees, null, 2));
        const workbook = new Workbook();
        workbook.calcProperties.fullCalcOnLoad = true;

        const paySummarySheet = workbook.addWorksheet('Pay Summary', { properties: { tabColor: { argb: 'FFC0000' } } });
        const summarySheet = workbook.addWorksheet('Summary', { properties: { tabColor: { argb: 'FFC0000' } } });

        this.generateSummarySheet(workbook, employees);
        this.generatePayrollSheet(workbook);
        this.generateEmployeeSheets(workbook, employees);

        this.resizeColumnsOfWorkbook(workbook);

        // this.printColumnStyle(summarySheet);
        // this.printRowStyle(summarySheet);
        workbook.xlsx.writeFile(filepath);
    }

    generatePayrollSheet = (workbook: Workbook) => {
        const summarySheet = workbook.getWorksheet('Summary');
        const paySummarySheet = workbook.getWorksheet('Pay Summary');

        this.fillPaySummaryHeaders(paySummarySheet)

        // tslint:disable-next-line: no-shadowed-variable
        for(let i = 3; i <= summarySheet.actualRowCount; i++) {
            // tslint:disable-next-line: no-shadowed-variable
            const row = paySummarySheet.getRow(i);
            for (let j = 1; j <= paySummarySheet.actualColumnCount; j++) {
                const column = paySummarySheet.getColumn(j);
                const key = column.key;
                const columnOffset = column.number;

                switch(key) {
                    // case 'ActualPay':
                    //     const startColumn = this.getColumnNumberFromHeader(paySummarySheet, 'Pay').letter;
                    //     const endColumn = this.getColumnNumberFromHeader(paySummarySheet, 'Bonus5').letter;

                    //     const formulaActualPay = `SUM(${startColumn}${i}:${endColumn}${i})`;
                    //     row.getCell(columnOffset).value = {
                    //         formula: formulaActualPay,
                    //         date1904: false
                    //     }
                    //     break;
                    // case 'TotalBonus':
                        // row.getCell(j).value = this.getTotalBonusFormula(paySummarySheet, i);
                        // break;
                    default:
                        const columnLetter = this.getColumnNumberFromHeader(summarySheet, key).letter;
                        const formula = `Summary!${columnLetter}${i}`
                        row.getCell(j).value = {
                            formula,
                            date1904: false
                        }

                        break;
                }
            }
        }

        const i = 2; // Totals row
        const row = paySummarySheet.getRow(i);
        for (let j = 1; j <= paySummarySheet.actualColumnCount; j++) {
            const column = paySummarySheet.getColumn(j);
            const key = column.key;
            const columnOffset = column.number;
            const columnLetter = column.letter;

            switch(key) {
                case 'Name':
                    row.getCell(columnOffset).value = 'TOTAL';
                    break;

                default:
                    const endRow = paySummarySheet.actualRowCount;
                    const formula = `SUM(${columnLetter}3:${columnLetter}${endRow})`
                    row.getCell(j).value = {
                        formula,
                        date1904: false
                    }

                    break;
            }
        }


        /* Style Summary Sheet */
        for (let x = 1; x <= 4; x++) {
            this.setColumnToColor(paySummarySheet, x, 'B7DEE8')
        }

        this.setRowToColor(paySummarySheet, 1, 'FFFF00')

        this.setRowToColor(paySummarySheet, 2, 'DA9694');
        /* END - Style Summary Sheet */

    }

    generateEmployeeSheets = (workbook: Workbook, employees: object) => {
        for (const employee in employees) {
            if (employees.hasOwnProperty(employee)) {
                if (employees[employee] && this.employeeList.List[employee]) {
                    employees[employee].HourlyRate = this.employeeList.List[employee].HourlyRate || 0;
                    employees[employee].StopRate = this.employeeList.List[employee].StopRate || 0;
                    employees[employee].DailyRate = this.employeeList.List[employee].DailyRate || 0;
                }

                let sheetName = employee
                if (employee === '') {
                    sheetName = 'blank'
                }
                const employeeSheet = workbook.addWorksheet(sheetName, { properties: { tabColor: { argb: 'C0C0C0' } } });
                this.fillEmployeeHeaders(employeeSheet);

                const employeeSummary = employees[employee]

                for (let i = 1; i <= employeeSheet.actualColumnCount; i++) {

                    const column = employeeSheet.getColumn(i);
                    const key = column.key;
                    const columnOffset = column.number;

                    let employeeRowOffset = 3;
                    for (const entry of employeeSummary.allEntries) {
                        // tslint:disable-next-line: no-shadowed-variable
                        const employeeRow = employeeSheet.getRow(employeeRowOffset);;
                        // console.log(employeeSummary['Name']);
                        // if (employeeSummary.Name.toLowerCase() === 'vinson,marcus lamar') {
                            // console.log(employeeRowOffset);
                            // console.log(entry);
                        // }
                        if (key === 'Name') {
                            // skip
                        } else if (key === 'HourlyRate') {
                            employeeRow.getCell(columnOffset).value = employees[employee].HourlyRate
                        } else if (key === 'HourlyPay') {
                            employeeRow.getCell(columnOffset).value = this.generateSameRowMultiplierFormulaForHourlyPay(employeeSheet, employeeRowOffset);
                            // employeeRow.getCell(columnOffset).value = this.generateSameRowMultiplierFormulaForHourlyPay(employeeSheet, 'HourlyRate', 'OnDutyHours', employeeRowOffset);
                            // employeeRow.getCell(columnOffset).value = this.calculateHourlyPay(employeeSummary.HourlyRate, entry.OnDutyHours)
                        } else if (key === 'StopRate') {
                            employeeRow.getCell(columnOffset).value = employees[employee].StopRate
                        } else if (key === 'StopPay') {
                            employeeRow.getCell(columnOffset).value = this.generateSameRowMultiplierFormula(employeeSheet, 'StopRate', 'AllStops', employeeRowOffset);
                            // employeeRow.getCell(columnOffset).value = this.calculateStopPay(employeeSummary.StopRate, entry.ActDelStps + entry.ActPUStps)
                        } else if (key === 'Pay') {
                            const StopPayColumn = this.getColumnNumberFromHeader(employeeSheet, 'StopPay').letter;
                            const HourlyPayColumn = this.getColumnNumberFromHeader(employeeSheet, 'HourlyPay').letter;

                            const formula = `MAX(${StopPayColumn}${employeeRowOffset},${HourlyPayColumn}${employeeRowOffset})`;
                            employeeRow.getCell(columnOffset).value = {
                                formula,
                                date1904: false
                            }
                        } else if (key === 'ActualPay') {
                            const startColumn = this.getColumnNumberFromHeader(employeeSheet, 'Pay').letter;
                            const endColumn = this.getColumnNumberFromHeader(employeeSheet, 'Bonus5').letter;

                            const formula = `SUM(${startColumn}${employeeRowOffset}:${endColumn}${employeeRowOffset})`;
                            employeeRow.getCell(columnOffset).value = {
                                formula,
                                date1904: false
                            }
                        } else if (key === 'OnRoadHours') {
                            employeeRow.getCell(columnOffset).value = this.convertHoursToNumber(entry[key])
                        } else if (key === 'OnDutyHours') {
                            employeeRow.getCell(columnOffset).value = this.convertHoursToNumber(entry[key])
                        } else if (key === 'PotDOTHrsViols') {
                            if (entry[key] === 'Y') {
                                this.setRowToColor(employeeSheet, employeeRowOffset, 'FFC000')
                            }
                            employeeRow.getCell(columnOffset).value = entry[key]
                        } else if (key === 'AllStops') {
                            employeeRow.getCell(columnOffset).value = entry.ActDelStps + entry.ActPUStps
                        } else {
                            employeeRow.getCell(columnOffset).value = entry[key];
                        }

                        employeeRowOffset++
                    }

                    // Totals on the top
                    employeeRowOffset = 2;
                    const employeeRow = employeeSheet.getRow(employeeRowOffset);
                    if (key === 'Name') {
                        employeeRow.getCell(columnOffset).value = 'Total';
                    } else if (key === 'Date') {
                        //
                    } else if (key === 'WAName') {
                        //
                    } else if (key === 'SvcAreaNumber') {
                        //
                    } else if (key === 'VehicleNumber') {
                        //
                    } else if (key === 'DriverHelperName') {
                        //
                    } else if (key === 'WANumber') {
                        //
                    } else if (key === 'HourlyRate') {
                        employeeRow.getCell(columnOffset).value = employees[employee].HourlyRate
                    } else if (key === 'HourlyPay') {
                        employeeRow.getCell(columnOffset).value = this.generateSameColumnSumFormula(3, employeeSheet.actualRowCount, column.letter);
                        // employeeRow.getCell(columnOffset).value = this.calculateHourlyPay(employeeSummary.HourlyRate, employeeSummary.OnDutyHours)
                        // employeeRow.getCell(columnOffset).value = employeeSummary.HourlyPay
                    } else if (key === 'StopRate') {
                        employeeRow.getCell(columnOffset).value = employees[employee].StopRate
                    } else if (key === 'StopPay') {
                        employeeRow.getCell(columnOffset).value = this.generateSameRowMultiplierFormula(employeeSheet, 'StopRate', 'AllStops', employeeRowOffset);
                        // employeeRow.getCell(columnOffset).value = employeeSummary.StopPay
                    } else if (key === 'Pay') {
                        employeeRow.getCell(columnOffset).value = this.generateSameColumnSumFormula(3, employeeSheet.actualRowCount, column.letter);
                    } else if (key === 'ActualPay') {
                        employeeRow.getCell(columnOffset).value = this.generateSameColumnSumFormula(3, employeeSheet.actualRowCount, column.letter);
                    } else if (key === 'OnRoadHours') {
                        employeeRow.getCell(columnOffset).value = this.generateSameColumnSumFormula(3, employeeSheet.actualRowCount, column.letter);
                    } else if (key === 'OnDutyHours') {
                        employeeRow.getCell(columnOffset).value = this.generateSameColumnSumFormula(3, employeeSheet.actualRowCount, column.letter);
                    } else if (key === 'RegularHours') {
                        employeeRow.getCell(columnOffset).value = this.generateRegularHoursFormula(employeeSheet, employeeRowOffset);
                    } else if (key === 'OTHours') {
                        employeeRow.getCell(columnOffset).value = this.generateOTHoursFormula(employeeSheet, employeeRowOffset);
                    } else if (key === 'PotDOTHrsViols') {
                        // if (employeeSummary.Totals[key] === 'Y') {
                        //     this.setRowToColor(employeeSheet, employeeRowOffset, 'FFC000')
                        // }
                        // employeeRow.getCell(columnOffset).value = employeeSummary.Totals[key]
                    } else {
                        employeeRow.getCell(columnOffset).value = this.generateSameColumnSumFormula(3, employeeSheet.actualRowCount, column.letter);
                    }

                }

                /* Style Employee Sheet */
                employeeSheet.views = [
                    // { state: 'frozen', xSplit: 17, ySplit: 2 }
                ]
                for (let x = 1; x <= 17; x++) {
                    employeeSheet.getColumn(x).style = { font: { bold: true }, numFmt: '#.00' }
                }
                for (let x = 1; x <= 17; x++) {
                    this.setColumnToColor(employeeSheet, x, 'B7DEE8');
                }

                for (let x = 1; x <= employeeSheet.actualColumnCount; x++) {
                    for (let y = 1; y <= 2; y++) {
                        employeeSheet.getRow(y).getCell(x).style = { font: { bold: true } }
                    }
                }

                this.setRowToColor(employeeSheet, 1, 'FFFF00');
                this.setRowToColor(employeeSheet, 2, 'B7DEE8');

                /* END - Style Employee Sheet */
            }
        }

    }
    generateSummarySheet = (workbook: Workbook, employees: object) => {
        // const summarySheet = workbook.addWorksheet('Summary', { properties: { tabColor: { argb: 'FFC0000' } } });
        const summarySheet = workbook.getWorksheet('Summary');
        this.fillSummaryHeaders(summarySheet)

        let rowOffset = 3;

        for (const employee in employees) {
            if (employees.hasOwnProperty(employee)) {
                if (employees[employee] && this.employeeList.List[employee]) {
                    employees[employee].HourlyRate = this.employeeList.List[employee].HourlyRate || 0;
                    employees[employee].StopRate = this.employeeList.List[employee].StopRate || 0;
                    employees[employee].DailyRate = this.employeeList.List[employee].DailyRate || 0;
                    employees[employee].PayType = this.employeeList.List[employee].PayType || '';
                }

                const employeeSummary = employees[employee]

                if (employee !== '') {
                    const row = summarySheet.getRow(rowOffset);
                    for (let i = 1; i <= summarySheet.actualColumnCount; i++) {
                        const column = summarySheet.getColumn(i);
                        // column.numFmt = '#.00';
                        const key = column.key;
                        const columnOffset = column.number;
                        if (key === 'Name') {
                            row.getCell(columnOffset).value = employee;
                        } else if (key === 'HourlyRate') {
                            row.getCell(columnOffset).value = employeeSummary.HourlyRate
                        } else if (key === 'HourlyPay') {
                            row.getCell(columnOffset).value = this.generateSameRowMultiplierFormulaForHourlyPay(summarySheet, rowOffset);
                            // row.getCell(columnOffset).value = this.generateSameRowMultiplierFormulaForHourlyPay(summarySheet, 'HourlyRate', 'OnDutyHours', rowOffset);
                            // row.getCell(columnOffset).value = this.calculateHourlyPay(employeeSummary.HourlyRate, employeeSummary.OnDutyHours)
                            // row.getCell(columnOffset).value = employeeSummary.HourlyPay
                        } else if (key === 'DailyRate') {
                            row.getCell(columnOffset).value = employeeSummary.DailyRate
                        } else if (key === 'DailyPay') {
                            row.getCell(columnOffset).value = this.generateSameRowMultiplierFormula(summarySheet, 'DailyRate', 'NumDays', rowOffset);
                            // row.getCell(columnOffset).value = employeeSummary.HourlyPay
                        } else if (key === 'StopRate') {
                            row.getCell(columnOffset).value = employeeSummary.StopRate
                        } else if (key === 'StopPay') {
                            row.getCell(columnOffset).value = this.generateSameRowMultiplierFormula(summarySheet, 'StopRate', 'AllStops', rowOffset);
                            // row.getCell(columnOffset).value = employeeSummary.StopPay
                        } else if (key === 'PayType') {
                            row.getCell(columnOffset).value = employees[employee].PayType
                        } else if (key === 'Pay') {
                            const StopPayColumn = this.getColumnNumberFromHeader(summarySheet, 'StopPay').letter;
                            const HourlyPayColumn = this.getColumnNumberFromHeader(summarySheet, 'HourlyPay').letter;
                            const DailyPayColumn = this.getColumnNumberFromHeader(summarySheet, 'DailyPay').letter;

                            console.log(`Employee = ${employee}, Paytype = ${employees[employee].PayType}`)
                            if(employees[employee] && employees[employee].PayType && employees[employee].PayType.toLowerCase() === 'daily') {
                                const formula = `=${DailyPayColumn}${rowOffset}`;
                                row.getCell(columnOffset).value = {
                                    formula,
                                    date1904: false
                                }
                            } else {
                                const formula = `MAX(${StopPayColumn}${rowOffset},${HourlyPayColumn}${rowOffset},${DailyPayColumn}${rowOffset})`;
                                row.getCell(columnOffset).value = {
                                    formula,
                                    date1904: false
                                }
                            }
                        } else if (key === 'ActualPay') {
                            const startColumn = this.getColumnNumberFromHeader(summarySheet, 'Pay').letter;
                            const endColumn = this.getColumnNumberFromHeader(summarySheet, 'Bonus5').letter;

                            const formula = `SUM(${startColumn}${rowOffset}:${endColumn}${rowOffset})`;
                            row.getCell(columnOffset).value = {
                                formula,
                                date1904: false
                            }
                        } else if (key === 'TotalBonus') {
                            if(employees[employee] && employees[employee].PayType && employees[employee].PayType.toLowerCase() === 'daily') {
                                const startColumn = this.getColumnNumberFromHeader(summarySheet, 'Bonus1').letter;
                                const endColumn = this.getColumnNumberFromHeader(summarySheet, 'Bonus5').letter;
                                const formula = `SUM(${startColumn}${rowOffset}:${endColumn}${rowOffset})`;
                                row.getCell(columnOffset).value = {
                                    formula,
                                    date1904: false
                                }
                            } else {
                                row.getCell(columnOffset).value = this.getTotalBonusFormula(summarySheet, rowOffset);
                            }
                        } else if (key === 'EffectiveHourlyRate') {
                            if(employees[employee] && employees[employee].PayType && employees[employee].PayType.toLowerCase() === 'daily') {
                                // Daily pay employee, no need for effective hours
                            } else {
                                const ActualPayColumn = this.getColumnNumberFromHeader(summarySheet, 'ActualPay').letter;
                                const NumHours = this.getColumnNumberFromHeader(summarySheet, 'OnDutyHours').letter;

                                const formula = `IF(${NumHours}${rowOffset} > 40, ${ActualPayColumn}${rowOffset}/(1.5*${NumHours}${rowOffset} - 20), ${ActualPayColumn}${rowOffset}/${NumHours}${rowOffset})`;
                                row.getCell(columnOffset).value = {
                                    formula,
                                    date1904: false
                                }
                            }
                        } else if (key === 'FinalPay') {
                            const actualPayColumn = this.getColumnNumberFromHeader(summarySheet, 'ActualPay').letter;
                            const discretionaryBonusColumn = this.getColumnNumberFromHeader(summarySheet, 'DiscretionaryBonus').letter;

                            const formula = `${actualPayColumn}${rowOffset}+${discretionaryBonusColumn}${rowOffset}`;
                            row.getCell(columnOffset).value = {
                                formula,
                                date1904: false
                            }
                        } else if (key === 'OnRoadHours') {
                            row.getCell(columnOffset).value = this.convertHoursToNumber(employeeSummary.Totals[key])
                        } else if (key === 'OnDutyHours') {
                            row.getCell(columnOffset).value = this.convertHoursToNumber(employeeSummary.Totals[key])
                        } else if (key === 'RegularHours') {
                            row.getCell(columnOffset).value = this.generateRegularHoursFormula(summarySheet, rowOffset);
                        } else if (key === 'OTHours') {
                            row.getCell(columnOffset).value = this.generateOTHoursFormula(summarySheet, rowOffset);
                        } else if (key === 'PotDOTHrsViols') {
                            if (employeeSummary.Totals[key] === 'Y') {
                                this.setRowToColor(summarySheet, rowOffset, 'FFC000')
                            }
                            row.getCell(columnOffset).value = employeeSummary.Totals[key];
                        } else {
                            row.getCell(columnOffset).value = employeeSummary.Totals[key];
                        }
                    }
                    rowOffset++;
                }
            }
        }

        /* Summary Sheet totals */
        rowOffset = 2;
        const cell = summarySheet.getRow(rowOffset).getCell(1).value = 'TOTALS';
        for (let columnNumber = 1; columnNumber <= summarySheet.actualColumnCount; columnNumber++) {
            const column = summarySheet.getColumn(columnNumber);
            const columnLetter = column.letter;
            const header = column.header;
            const maxRow = summarySheet.actualRowCount;
            const val = {
                formula: `SUM(${columnLetter}3:${columnLetter}${maxRow})`,
                date1904: false
            }
            switch (header) {
                case 'AllStops':
                case 'AllStops':
                case 'Pay':
                case 'Bonus1':
                case 'Bonus2':
                case 'Bonus3':
                case 'Bonus4':
                case 'Bonus5':
                case 'ActualPay':
                case 'HourlyPay':
                case 'OnDutyHours':
                case 'RegularHours':
                case 'OTHours':
                case 'VScanPkgs':
                case 'DelStps':
                case 'DIFF':
                case 'ActDelStps':
                case 'ActDelPkgs':
                case 'ActPUStps':
                case 'ActPUPkgs':
                case 'ILSImpactPkgs':
                case 'NondelvdStps':
                case 'Code85':
                case 'AllStatusCodePkgs':
                case 'PLML':
                case 'DNA':
                case 'SndAgn':
                case 'Excs':
                case 'VSAvsSTARDIFF':
                case 'OnRoadHours':
                case 'OnDutyHours':
                case 'PotMissPUs':
                case 'ELPUs':
                case 'ReqSig':
                case 'DateCertain':
                case 'Evening':
                case 'Appt':
                    summarySheet.getRow(rowOffset).getCell(columnNumber).value = val;
                    break;
                case 'PotDOTHrsViols':
                    summarySheet.getRow(rowOffset).getCell(columnNumber).value = {
                        formula: `=ISNUMBER(SEARCH("Y",${columnLetter}3:${columnLetter}${maxRow}))`,
                        date1904: false
                    };
                    break;
            }
        }
        /* Style Summary Sheet */
        summarySheet.views = [
            // { state: 'frozen', xSplit: 17, ySplit: 2 }
        ]
        for (let x = 1; x <= 23; x++) {
            this.setColumnToColor(summarySheet, x, 'B7DEE8')
        }

        this.setRowToColor(summarySheet, 1, 'FFFF00')

        this.setRowToColor(summarySheet, 2, 'DA9694');
        /* END - Style Summary Sheet */
    }

    resizeColumnsOfWorkbook = (workbook: Workbook) => {
        workbook.eachSheet((sheet, id) => {
            sheet.columns.forEach((column) => {
                if (column.key === 'Date') {
                    column.width = '00/00/0000'.length
                } else {
                    column.eachCell({ includeEmpty: false }, (cell) => {
                        if (cell.value) {
                            const columnLength = cell.value.toString().length;
                            // console.log('cell.value.toString()', `'${cell.value.toString()}'`)
                            if (column.width < columnLength) {
                                column.width = columnLength
                            }
                        }
                    })
                }
            });
        })
    }

    generateSameRowMultiplierFormulaForHourlyPay = (sheet: Worksheet, rowNumber): CellFormulaValue => {
    // generateSameRowMultiplierFormulaForHourlyPay = (sheet: Worksheet, column1Header, column2Header, rowNumber): CellFormulaValue => {
        // generateSameRowMultiplierFormula(sheet, 'HourlyRate', 'OnDutyHours', rowOffset);
        // const column1 = this.getColumnNumberFromHeader(sheet, column1Header);
        // const column2 = this.getColumnNumberFromHeader(sheet, column2Header);

        // const numHours = sheet.getRow(rowNumber).getCell(column2.number).value;

        // if (numHours > 40) {
        //     if (column1 !== null && column2 !== null) {
        //         const column1Letter = column1.letter;
        //         const column2Letter = column2.letter;
        //         const formula = `${column2Letter}${rowNumber}*${column1Letter}${rowNumber} + ((${column2Letter}${rowNumber} - 40)*${column1Letter}${rowNumber})`;
        //         return { formula, date1904: false };
        //     }
        // } else {
        //     if (column1 !== null && column2 !== null) {
        //         const column1Letter = column1.letter;
        //         const column2Letter = column2.letter;
        //         const formula = column1Letter + rowNumber + '*' + column2Letter + rowNumber;
        //         return { formula, date1904: false };
        //     }
        // }

        // return null;

        const RegularHoursColumn = this.getColumnNumberFromHeader(sheet, 'RegularHours');
        const RegularHoursColumnLetter = RegularHoursColumn.letter;

        const OTHoursColumn = this.getColumnNumberFromHeader(sheet, 'OTHours');
        const OTHoursColumnLetter = OTHoursColumn.letter;

        const HourlyRateColumn = this.getColumnNumberFromHeader(sheet, 'HourlyRate');
        const HourlyRateColumnLetter = HourlyRateColumn.letter;

        const formula = `(${RegularHoursColumnLetter}${rowNumber}*${HourlyRateColumnLetter}${rowNumber}) + (${OTHoursColumnLetter}${rowNumber}*${HourlyRateColumnLetter}${rowNumber}*1.5)`
        return { formula, date1904: false };
    }

    generateSameRowMultiplierFormula = (sheet: Worksheet, column1Header, column2Header, rowNumber): CellFormulaValue => {
        // generateSameRowMultiplierFormula(sheet, 'HourlyRate', 'OnDutyHours', rowOffset);
        const column1 = this.getColumnNumberFromHeader(sheet, column1Header);
        const column2 = this.getColumnNumberFromHeader(sheet, column2Header);

        if (column1 !== null && column2 !== null) {
            const column1Letter = column1.letter;
            const column2Letter = column2.letter;
            const formula = column1Letter + rowNumber + '*' + column2Letter + rowNumber;
            return { formula, date1904: false };
        }

        return null;
    }

    generateRegularHoursFormula = (sheet: Worksheet, rowNumber:number): CellFormulaValue => {
        const onDutyHoursColumn = this.getColumnNumberFromHeader(sheet, 'OnDutyHours');
        const onDutyHoursColumnLetter = onDutyHoursColumn.letter;
        const formula = `IF(${onDutyHoursColumnLetter}${rowNumber} > 40, 40, ${onDutyHoursColumnLetter}${rowNumber})`
        return { formula, date1904: false };
    }

    generateOTHoursFormula = (sheet: Worksheet, rowNumber:number): CellFormulaValue => {
        const onDutyHoursColumn = this.getColumnNumberFromHeader(sheet, 'OnDutyHours');
        const onDutyHoursColumnLetter = onDutyHoursColumn.letter;

        const formula = `IF(${onDutyHoursColumnLetter}${rowNumber} > 40, ${onDutyHoursColumnLetter}${rowNumber} - 40, 0)`
        return { formula, date1904: false };
    }


    getTotalBonusFormula = (sheet: Worksheet, rowNumber:number): CellFormulaValue => {
        const HourlyPayColumn = this.getColumnNumberFromHeader(sheet, 'HourlyPay');
        const HourlyPayColumnLetter = HourlyPayColumn.letter;

        const ActualPayColumn = this.getColumnNumberFromHeader(sheet, 'ActualPay');
        const ActualPayColumnLetter = ActualPayColumn.letter;

        const formula = `IF(${ActualPayColumnLetter}${rowNumber} > ${HourlyPayColumnLetter}${rowNumber}, ${ActualPayColumnLetter}${rowNumber} - ${HourlyPayColumnLetter}${rowNumber}, 0)`
        return { formula, date1904: false };
    }
    generateSameColumnSumFormula = (fromRowNumber: number, toRowNumber: number, columnLetter: string): CellFormulaValue => {
        const formula = `SUM(${columnLetter}${fromRowNumber}:${columnLetter}${toRowNumber})`;
        return { formula, date1904: false };
    }

    convertHoursToNumber = (hours: Hours): number => {
        const numOfPartialHours = hours.minutes / 60;
        return Number((hours.hours + numOfPartialHours).toFixed(2))
    }

    getColumnNumberFromHeader = (sheet: Worksheet, header: string): Partial<Column> => {
        let column: Partial<Column> = null;
        for (let i = 1; i <= sheet.actualColumnCount; i++) {
            const c = sheet.getColumn(i);
            if (c.header === header) {
                column = c;
                break;
            }
        }
        return column;
    }

    printColumnStyle = (sheet: Worksheet) => {
        for (let i = 1; i <= sheet.actualColumnCount; i++) {
            const column = sheet.getColumn(i);
            // console.log(column.letter, column.style);
        }
    }

    printRowStyle = (sheet: Worksheet) => {
        for (let i = 1; i <= sheet.actualRowCount; i++) {
            const row = sheet.getRow(i);
            // console.log(row.number, row.numFmt);
        }
    }

    fillPaySummaryHeaders = (sheet: Worksheet) => {
        sheet.columns = [
            { header: 'Name', key: 'Name' },
            // { header: 'PayType', key: 'PayType' },
            // { header: 'EffectiveHourlyRate', key: 'EffectiveHourlyRate', style: { numFmt: '#.00' } },
            // { header: 'RegularHours', key: 'RegularHours', style: { numFmt: '#.00' } },
            // { header: 'OTHours', key: 'OTHours', style: { numFmt: '#.00' } },
            // { header: 'DailyRate', key: 'DailyRate', style: { numFmt: '#.00' } },
            // { header: 'NumDays', key: 'NumDays', style: { numFmt: '#.00' } },
            // { header: 'HourlyPay', key: 'HourlyPay', style: { numFmt: '#.00' } },
            // { header: 'DailyPay', key: 'DailyPay', style: { numFmt: '#.00' } },
            // { header: 'StopPay', key: 'StopPay', style: { numFmt: '#.00' } },
            // { header: 'Pay', key: 'Pay', style: { numFmt: '#.00' } },
            // { header: 'Bonus1', key: 'Bonus1', style: { numFmt: '#.00' } },
            // { header: 'Bonus2', key: 'Bonus2', style: { numFmt: '#.00' } },
            // { header: 'Bonus3', key: 'Bonus3', style: { numFmt: '#.00' } },
            // { header: 'Bonus4', key: 'Bonus4', style: { numFmt: '#.00' } },
            // { header: 'Bonus5', key: 'Bonus5', style: { numFmt: '#.00' } },
            // { header: 'TotalBonus', key: 'TotalBonus', style: { numFmt: '#.00' } },
            { header: 'ActualPay', key: 'ActualPay', style: { numFmt: '#.00' } },
            { header: 'DiscretionaryBonus', key: 'DiscretionaryBonus', style: { numFmt: '#.00' } },
            { header: 'FinalPay', key: 'FinalPay', style: { numFmt: '#.00' } },
        ];
    }


    fillSummaryHeaders = (sheet: Worksheet) => {
        sheet.columns = [
            { header: 'Name', key: 'Name' },
            { header: 'PayType', key: 'PayType' },
            { header: 'HourlyRate', key: 'HourlyRate', style: { numFmt: '#.00' } },
            { header: 'EffectiveHourlyRate', key: 'EffectiveHourlyRate', style: { numFmt: '#.00' } },
            { header: 'OnDutyHours', key: 'OnDutyHours', style: { numFmt: '#.00' } },
            { header: 'RegularHours', key: 'RegularHours', style: { numFmt: '#.00' } },
            { header: 'OTHours', key: 'OTHours', style: { numFmt: '#.00' } },
            { header: 'DailyRate', key: 'DailyRate', style: { numFmt: '#.00' } },
            { header: 'NumDays', key: 'NumDays', style: { numFmt: '#.00' } },
            { header: 'StopRate', key: 'StopRate', style: { numFmt: '#.00' } },
            { header: 'AllStops', key: 'AllStops' },
            { header: 'HourlyPay', key: 'HourlyPay', style: { numFmt: '#.00' } },
            { header: 'DailyPay', key: 'DailyPay', style: { numFmt: '#.00' } },
            { header: 'StopPay', key: 'StopPay', style: { numFmt: '#.00' } },
            { header: 'Pay', key: 'Pay', style: { numFmt: '#.00' } },
            { header: 'Bonus1', key: 'Bonus1', style: { numFmt: '#.00' } },
            { header: 'Bonus2', key: 'Bonus2', style: { numFmt: '#.00' } },
            { header: 'Bonus3', key: 'Bonus3', style: { numFmt: '#.00' } },
            { header: 'Bonus4', key: 'Bonus4', style: { numFmt: '#.00' } },
            { header: 'Bonus5', key: 'Bonus5', style: { numFmt: '#.00' } },
            { header: 'ActualPay', key: 'ActualPay', style: { numFmt: '#.00' } },
            // { header: 'TotalBonus', key: 'TotalBonus', style: { numFmt: '#.00' } },
            { header: 'DiscretionaryBonus', key: 'DiscretionaryBonus', style: { numFmt: '#.00' } },
            { header: 'FinalPay', key: 'FinalPay', style: { numFmt: '#.00' } },

            { header: 'VScanPkgs', key: 'VScanPkgs' },
            { header: 'DelStps', key: 'DelStps' },
            { header: 'PUStps', key: 'PUStps' },
            { header: 'DIFF', key: 'DIFF' },
            { header: 'ActDelStps', key: 'ActDelStps' },
            { header: 'ActDelPkgs', key: 'ActDelPkgs' },
            { header: 'ActPUStps', key: 'ActPUStps' },
            { header: 'ActPUPkgs', key: 'ActPUPkgs' },
            { header: 'ILSImpactPkgs', key: 'ILSImpactPkgs' },
            { header: 'NondelvdStps', key: 'NondelvdStps' },
            { header: 'Code85', key: 'Code85' },
            { header: 'AllStatusCodePkgs', key: 'AllStatusCodePkgs' },
            { header: 'PLML', key: 'PLML' },
            { header: 'DNA', key: 'DNA' },
            { header: 'SndAgn', key: 'SndAgn' },
            { header: 'Excs', key: 'Excs' },
            { header: 'VSAvsSTARDIFF', key: 'VSAvsSTARDIFF' },
            { header: 'Miles', key: 'Miles' },
            { header: 'OnRoadHours', key: 'OnRoadHours', style: { numFmt: '#.00' } },
            { header: 'OnDutyHours', key: 'OnDutyHours', style: { numFmt: '#.00' } },
            { header: 'PotDOTHrsViols', key: 'PotDOTHrsViols' },
            { header: 'PotMissPUs', key: 'PotMissPUs' },
            { header: 'ELPUs', key: 'ELPUs' },
            { header: 'ReqSig', key: 'ReqSig' },
            { header: 'DateCertain', key: 'DateCertain' },
            { header: 'Evening', key: 'Evening' },
            { header: 'Appt', key: 'Appt' },
        ];
    }
    fillEmployeeHeaders = (sheet: Worksheet) => {
        sheet.columns = [
            { header: 'Name', key: 'Name' },
            { header: 'HourlyRate', key: 'HourlyRate', style: { numFmt: '#.00' } },
            { header: 'OnDutyHours', key: 'OnDutyHours', style: { numFmt: '#.00' } },
            { header: 'RegularHours', key: 'RegularHours', style: { numFmt: '#.00' } },
            { header: 'OTHours', key: 'OTHours', style: { numFmt: '#.00' } },
            { header: 'DailyRate', key: 'DailyRate', style: { numFmt: '#.00' } },
            { header: 'NumDays', key: 'NumDays', style: { numFmt: '#.00' } },
            { header: 'StopRate', key: 'StopRate', style: { numFmt: '#.00' } },
            { header: 'AllStops', key: 'AllStops' },
            { header: 'HourlyPay', key: 'HourlyPay', style: { numFmt: '#.00' } },
            { header: 'DailyPay', key: 'DailyPay', style: { numFmt: '#.00' } },
            { header: 'StopPay', key: 'StopPay', style: { numFmt: '#.00' } },
            { header: 'Pay', key: 'Pay', style: { numFmt: '#.00' } },
            { header: 'Bonus1', key: 'Bonus1', style: { numFmt: '#.00' } },
            { header: 'Bonus2', key: 'Bonus2', style: { numFmt: '#.00' } },
            { header: 'Bonus3', key: 'Bonus3', style: { numFmt: '#.00' } },
            { header: 'Bonus4', key: 'Bonus4', style: { numFmt: '#.00' } },
            { header: 'Bonus5', key: 'Bonus5', style: { numFmt: '#.00' } },
            { header: 'ActualPay', key: 'ActualPay', style: { numFmt: '#.00' } },

            { header: 'Date', key: 'Date' },
            { header: 'WAName', key: 'WAName' },
            { header: 'SvcAreaNumber', key: 'SvcAreaNumber' },
            { header: 'VehicleNumber', key: 'VehicleNumber' },
            { header: 'DriverHelperName', key: 'DriverHelperName' },
            { header: 'WANumber', key: 'WANumber' },
            { header: 'MDSG', key: 'MDSG' },
            { header: 'DST', key: 'VScanDSTPkgs' },
            { header: 'VScanPkgs', key: 'VScanPkgs' },
            { header: 'DelStps', key: 'DelStps' },
            { header: 'PUStps', key: 'PUStps' },
            { header: 'DIFF', key: 'DIFF' },
            { header: 'ActDelStps', key: 'ActDelStps' },
            { header: 'ActDelPkgs', key: 'ActDelPkgs' },
            { header: 'ActPUStps', key: 'ActPUStps' },
            { header: 'ActPUPkgs', key: 'ActPUPkgs' },
            { header: 'ILSImpactPkgs', key: 'ILSImpactPkgs' },
            { header: 'NondelvdStps', key: 'NondelvdStps' },
            { header: 'Code85', key: 'Code85' },
            { header: 'AllStatusCodePkgs', key: 'AllStatusCodePkgs' },
            { header: 'PLML', key: 'PLML' },
            { header: 'DNA', key: 'DNA' },
            { header: 'SndAgn', key: 'SndAgn' },
            { header: 'Excs', key: 'Excs' },
            { header: 'VSAvsSTARDIFF', key: 'VSAvsSTARDIFF' },
            { header: 'Miles', key: 'Miles' },
            { header: 'OnRoadHours', key: 'OnRoadHours', style: { numFmt: '#.00' } },
            { header: 'OnDutyHours', key: 'OnDutyHours', style: { numFmt: '#.00' } },
            { header: 'PotDOTHrsViols', key: 'PotDOTHrsViols' },
            { header: 'PotMissPUs', key: 'PotMissPUs' },
            { header: 'ELPUs', key: 'ELPUs' },
            { header: 'ReqSig', key: 'ReqSig' },
            { header: 'DateCertain', key: 'DateCertain' },
            { header: 'Evening', key: 'Evening' },
            { header: 'Appt', key: 'Appt' },
        ];
    }

    setRowToColor = (sheet: Worksheet, row: number, color: string) => {
        for (let y = 1; y <= sheet.actualColumnCount; y++) {
            const style = sheet.getRow(row).getCell(y).style;
            style.font = style.font || {};
            style.font.bold = true;
            sheet.getRow(row).getCell(y).style = style
            sheet.getRow(row).getCell(y).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: color }
            }
        }
    }

    setColumnToColor = (sheet: Worksheet, column: number, color: string) => {
        for (let y = 2; y <= sheet.actualRowCount; y++) {
            const style = sheet.getRow(y).getCell(column).style;
            style.font = style.font || {};
            style.font.bold = true;
            sheet.getRow(y).getCell(column).style = style
            sheet.getRow(y).getCell(column).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: color }
            }
        }
    }
}