import { Workbook, Worksheet, Row, ValueType, CellValue } from 'exceljs';


export class Employee {
    Name:string
    PayType:string
    HourlyRate:number
    DailyRate:number
    MonthlyRate:number
    StopRate:number
    BulkStopBonus:number
    Filepath:string

    columnList:string[] = []

    List:Map<string, Employee>;

    constructor (filepath:string = '') {
        this.Filepath = filepath
    }

    init = async () => {
        if(this.Filepath === '') {
            throw new Error('Filepath not provided')
        }
        await this.loadFromFile(this.Filepath);
    }

    loadFromFile = async (filepath:string) => {
        const workbook = new Workbook();
        await workbook.xlsx.readFile(filepath);

        const employeewWorksheet = workbook.getWorksheet(1);

        this.scanColumns(employeewWorksheet);

        this.List = this.populateEmployeeList(employeewWorksheet);
    }

    populateEmployeeList = (worksheet:Worksheet):Map<string, Employee> => {
        const employees:Map<string, Employee> = new Map<string, Employee>();

        for(let i = 1; i <= worksheet.actualRowCount; i++) {
            const employee = new Employee();
            const row = worksheet.getRow(i);

            for(let j = 1; j <= worksheet.actualColumnCount; j++) {
                const cell = row.getCell(j);
                let value = cell.value;
                const columnType = this.columnList[j-1];

                if(columnType !== 'Name' && columnType !== 'PayType') {
                    value = value || 0;
                    employee[columnType] = parseFloat(value.toString()) || 0;
                } else {
                    if(value !== null) {
                        employee[columnType] = value.toString();
                    }
                }

                // if(columnType === 'DailyRate') {
                //     console.log(employee[columnType]);
                // }
                // if(typeof(employee[columnType]) === 'number') {
                    // employee[columnType] = parseFloat(value.toString()) || 0;
                // } else {
                    // employee[columnType] = value;
                // }
            }
            employees[employee.Name] = employee;
        }
        return employees;
    }

    scanColumns = (worksheet:Worksheet) => {
        for(let i = 1; i <= worksheet.actualColumnCount; i++) {
            const value = worksheet.getRow(1).getCell(i).value;
            let name = "";
            // console.log(value);
            switch(value) {
                case 'Name': name = 'Name';
                             break;

                case 'Pay Type': name = 'PayType';
                             break;

                case 'Hourly Rate': name = 'HourlyRate';
                             break;

                case 'Daily Rate': name = 'DailyRate';
                             break;

                case 'Monthly Rate': name = 'MonthlyRate';
                             break;

                case 'Stop Rate': name = 'StopRate';
                             break;

                case 'Bulk Stops Bonus': name = 'BulkStopsBonus';
                             break;
            }
            this.columnList[i-1] = name;
        }
    }

}