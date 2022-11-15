export interface Hours {
    hours:number
    minutes:number
}

export class Entry {
    Date:Date;
    WAName:string;
    SvcAreaNumber:number;
    VehicleNumber:number;
    DriverHelperName:string;
    DriverName:string;
    HelperName:string;
    WANumber:number;
    MDSG:string;
    DST:number;
    VScanPkgs:number;
    DelStps:number;
    PUStps:number;
    DIFF:number;
    ActDelStps:number;
    ActDelPkgs:number;
    ActPUStps:number;
    ActPUPkgs:number;
    ILSPercent:number;
    ILSImpactPkgs:number;
    NondelvdStps:number;
    Code85:number;
    AllStatusCodePkgs:number;
    PLML:number;
    DNA:number;
    SndAgn:number;
    Excs:number;
    VSAvsSTARDIFF:number;
    ReturnsScansPercent:number;
    Miles:number;
    OnRoadHours:Hours;
    OnDutyHours:Hours;
    PotDOTHrsViols:string;
    NextAvailOnDuty:Hours;
    PotMissPUs:number;
    ELPUs:number;
    ReqSig:number;
    DateCertain:number;
    Evening:number;
    Appt:number;

}
