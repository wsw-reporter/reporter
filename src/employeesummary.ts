import { Totals } from "./totals";
import { Entry } from "./entry";

export class EmployeeSummary {
    allEntries:Entry[];
    Totals:Totals;
    HourlyPay:number;
    DailyPay:number;
    StopPay:number;
    BulkStopBonus:number;
    Bonus:number;
    HourlyRate:number;
    DailyRate:number;
    StopRate:number;
    PayType:string;
    Name:string;


    constructor() {
        this.Totals = {
            VScanPkgs: 0,
            DelStps: 0,
            PUStps: 0,
            DIFF: 0,
            ActDelStps: 0,
            ActDelPkgs: 0,
            ActPUStps: 0,
            ActPUPkgs: 0,
            ILSImpactPkgs: 0,
            NondelvdStps: 0,
            Code85: 0,
            AllStatusCodePkgs: 0,
            PLML: 0,
            DNA: 0,
            SndAgn: 0,
            Excs: 0,
            VSAvsSTARDIFF: 0,
            Miles: 0,
            OnRoadHours: { hours: 0, minutes: 0 },
            OnDutyHours: { hours: 0, minutes: 0 },
            PotDOTHrsViols: { hours: 0, minutes: 0 },
            PotMissPUs: 0,
            ELPUs: 0,
            ReqSig: 0,
            DateCertain: 0,
            Evening: 0,
            Appt: 0,
            AllStops: 0,
            NumDays: 0
        }
    }

    addEntry = function(entry:Entry) {
        this.allEntries = this.allEntries || [];
        this.allEntries.push(entry);

        this.Totals.VScanPkgs += entry.VScanPkgs;
        this.Totals.DelStps += entry.DelStps;
        this.Totals.PUStps += entry.PUStps
        this.Totals.DIFF += entry.DIFF;
        this.Totals.ActDelStps += entry.ActDelStps;
        this.Totals.ActDelPkgs += entry.ActDelPkgs;
        this.Totals.ActPUStps += entry.ActPUStps;
        this.Totals.ActPUPkgs += entry.ActPUPkgs;
        this.Totals.ILSImpactPkgs += entry.ILSImpactPkgs;
        this.Totals.NondelvdStps += entry.NondelvdStps;
        this.Totals.Code85 += entry.Code85;
        this.Totals.AllStatusCodePkgs += entry.AllStatusCodePkgs;
        this.Totals.PLML += entry.PLML;
        this.Totals.DNA += entry.DNA;
        this.Totals.SndAgn += entry.SndAgn;
        this.Totals.Excs += entry.Excs;
        this.Totals.VSAvsSTARDIFF += entry.VSAvsSTARDIFF;
        this.Totals.Miles += entry.Miles
        this.Totals.OnRoadHours.hours += entry.OnRoadHours.hours;
        this.Totals.OnRoadHours.minutes += entry.OnRoadHours.minutes;
        this.Totals.OnDutyHours.hours += entry.OnDutyHours.hours;
        this.Totals.OnDutyHours.minutes += entry.OnDutyHours.minutes;
        this.Totals.PotDOTHrsViols = (entry.PotDOTHrsViols === 'Y')?'Y':'';
        this.Totals.PotMissPUs += entry.PotMissPUs;
        this.Totals.ELPUs += entry.ELPUs;
        this.Totals.ReqSig += entry.ReqSig;
        this.Totals.DateCertain += entry.DateCertain;
        this.Totals.Evening += entry.Evening;
        this.Totals.Appt += entry.Appt;

        this.Totals.AllStops += entry.ActDelStps + entry.ActPUStps
        this.Totals.NumDays = this.getUniqueDaysCount()
    }

    getUniqueDaysCount = ():number => {
        const days:Map<string,{}> = new Map<string,{}>();
        for(const entry of this.allEntries) {
            const dateString = entry.Date.getFullYear().toString() + entry.Date.getMonth().toString() + entry.Date.getDate().toString();
            days[dateString] = {}
        }

        return Object.keys(days).length;
    }
}
