﻿//------------------------------------------------------------------------------
// <auto-generated>
//    Este código se generó a partir de una plantilla.
//
//    Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//    Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CentralDashboard.Models.EntiCorporativa
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    using System.Data.Objects;
    using System.Data.Objects.DataClasses;
    using System.Linq;
    
    public partial class Entities : DbContext
    {
        public Entities()
            : base("name=Entities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public DbSet<USR_PaginaSitioWeb> USR_PaginaSitioWeb { get; set; }
        public DbSet<USR_PermisoSitioWeb> USR_PermisoSitioWeb { get; set; }
    
        public virtual ObjectResult<RPT_DiarioHospitalizacion_Result> RPT_DiarioHospitalizacion()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<RPT_DiarioHospitalizacion_Result>("RPT_DiarioHospitalizacion");
        }
    }
}
