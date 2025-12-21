using System;
using System.Collections.Generic;

namespace XSLXtoCSV.Data.UPM_System;

public partial class ProductionAchievement
{
    public Guid Id { get; set; }

    public bool Active { get; set; }

    public DateTime CreateDate { get; set; }

    public string CreateBy { get; set; } = null!;

    public DateTime ProductionDate { get; set; }

    public string Supervisor { get; set; } = null!;

    public string Leader { get; set; } = null!;

    public string Shift { get; set; } = null!;

    public string PartNumberName { get; set; } = null!;

    public Guid? PartNumberId { get; set; }

    public float WorkingTime { get; set; }

    public float ProductionObjetive { get; set; }

    public float ProductionReal { get; set; }

    public string Area { get; set; } = null!;
}
