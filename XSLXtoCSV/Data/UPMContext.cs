using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using XSLXtoCSV.Data.UPM_System;

namespace XSLXtoCSV.Data;

public partial class UPMContext : DbContext
{
    public UPMContext()
    {
    }

    public UPMContext(DbContextOptions<UPMContext> options)
        : base(options)
    {
    }

    public virtual DbSet<OperationalEfficiency> OperationalEfficiencies { get; set; }

    public virtual DbSet<ProductionAchievement> ProductionAchievements { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see https://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=UPMAP04\\UPMDATA;Database=UPMWEB;User Id=upmUser;Password=Flg6petoa3z3UEyyglDA;TrustServerCertificate=True;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.UseCollation("Modern_Spanish_CI_AS");

        modelBuilder.Entity<OperationalEfficiency>(entity =>
        {
            entity.ToTable("OperationalEfficiency", "upm_temporal");

            entity.Property(e => e.Id).HasDefaultValueSql("(newid())");
            entity.Property(e => e.Hp).HasColumnName("HP");
            entity.Property(e => e.Leader).HasMaxLength(100);
            entity.Property(e => e.PartNumberName).HasMaxLength(200);
            entity.Property(e => e.Shift).HasMaxLength(50);
            entity.Property(e => e.Supervisor).HasMaxLength(100);
        });

        modelBuilder.Entity<ProductionAchievement>(entity =>
        {
            entity.ToTable("ProductionAchievement", "upm_temporal");

            entity.Property(e => e.Id).HasDefaultValueSql("(newid())");
            entity.Property(e => e.Area).HasDefaultValue("");
            entity.Property(e => e.Leader).HasMaxLength(100);
            entity.Property(e => e.PartNumberName).HasMaxLength(200);
            entity.Property(e => e.Shift).HasMaxLength(50);
            entity.Property(e => e.Supervisor).HasMaxLength(100);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
