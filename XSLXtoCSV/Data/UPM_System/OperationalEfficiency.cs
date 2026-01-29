using System;
using System.Collections.Generic;

namespace XSLXtoCSV.Data.UPM_System;

public partial class OperationalEfficiency
{
    public Guid Id { get; set; }

    public bool Active { get; set; }

    public DateTime CreateDate { get; set; }

    public string CreateBy { get; set; } = null!;

    public DateTime ProductionDate { get; set; }

    public string Area { get; set; } = null!;

    public string Supervisor { get; set; } = null!;

    public string Leader { get; set; } = null!;

    public string Shift { get; set; } = null!;

    public string PartNumberName { get; set; } = null!;

    public float Hp { get; set; }

    public float Neck { get; set; }

    public Guid? PartNumberId { get; set; }

    public float RealTime { get; set; }

    public float OperativityPercent { get; set; }

    public float PriductionReal { get; set; }

    public float TotalTime { get; set; }

    public float ProgramabeDowntimeTime { get; set; }

    public float RealWorkingTime { get; set; }

    public float NetoWorkingTime { get; set; }

    public float NetoProduictiveTime { get; set; }

    public float TotalDowntime { get; set; }

    public float NoProgramabeDowntimeTime { get; set; }

    public float NoReportedTime { get; set; }

    public float DowntimePercent { get; set; }

    public float NoProgramableDowntimePercent { get; set; }

    public float ProgramableDowntimePercent { get; set; }

    public float Aprov { get; set; }

    public float Junta { get; set; }

    public float Pilotaje { get; set; }

    public float SpmReal { get; set; }

    public float SpmSet { get; set; }

    public float StSpmSet { get; set; }

    public float Stroke { get; set; }

    public float Tt { get; set; }

    public float Ttt { get; set; }

    public string Jefe { get; set; } = null!;

    public string Managment { get; set; } = null!;
}
