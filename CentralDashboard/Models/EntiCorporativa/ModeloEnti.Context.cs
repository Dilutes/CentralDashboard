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
    
        public virtual ObjectResult<REM_DatosBase_Result> REM_DatosBase(Nullable<int> mes, Nullable<int> anio)
        {
            var mesParameter = mes.HasValue ?
                new ObjectParameter("mes", mes) :
                new ObjectParameter("mes", typeof(int));
    
            var anioParameter = anio.HasValue ?
                new ObjectParameter("anio", anio) :
                new ObjectParameter("anio", typeof(int));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<REM_DatosBase_Result>("REM_DatosBase", mesParameter, anioParameter);
        }
    
        public virtual ObjectResult<Nullable<int>> REM_GetAños()
        {
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<Nullable<int>>("REM_GetAños");
        }
    
        public virtual ObjectResult<Nullable<int>> REM_GetMeses(Nullable<int> anio)
        {
            var anioParameter = anio.HasValue ?
                new ObjectParameter("anio", anio) :
                new ObjectParameter("anio", typeof(int));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<Nullable<int>>("REM_GetMeses", anioParameter);
        }
    
        public virtual ObjectResult<REM_SeccionA_Result> REM_SeccionA(Nullable<int> mes, Nullable<int> anio)
        {
            var mesParameter = mes.HasValue ?
                new ObjectParameter("mes", mes) :
                new ObjectParameter("mes", typeof(int));
    
            var anioParameter = anio.HasValue ?
                new ObjectParameter("anio", anio) :
                new ObjectParameter("anio", typeof(int));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<REM_SeccionA_Result>("REM_SeccionA", mesParameter, anioParameter);
        }
    
        public virtual ObjectResult<REM_SeccionB_Result> REM_SeccionB(Nullable<int> mes, Nullable<int> anio)
        {
            var mesParameter = mes.HasValue ?
                new ObjectParameter("mes", mes) :
                new ObjectParameter("mes", typeof(int));
    
            var anioParameter = anio.HasValue ?
                new ObjectParameter("anio", anio) :
                new ObjectParameter("anio", typeof(int));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<REM_SeccionB_Result>("REM_SeccionB", mesParameter, anioParameter);
        }
    
        public virtual ObjectResult<REM_SeccionD_Result> REM_SeccionD(Nullable<int> mes, Nullable<int> anio)
        {
            var mesParameter = mes.HasValue ?
                new ObjectParameter("mes", mes) :
                new ObjectParameter("mes", typeof(int));
    
            var anioParameter = anio.HasValue ?
                new ObjectParameter("anio", anio) :
                new ObjectParameter("anio", typeof(int));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<REM_SeccionD_Result>("REM_SeccionD", mesParameter, anioParameter);
        }
    
        public virtual ObjectResult<REM_SeccionF_Result> REM_SeccionF(Nullable<int> mes, Nullable<int> anio)
        {
            var mesParameter = mes.HasValue ?
                new ObjectParameter("mes", mes) :
                new ObjectParameter("mes", typeof(int));
    
            var anioParameter = anio.HasValue ?
                new ObjectParameter("anio", anio) :
                new ObjectParameter("anio", typeof(int));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<REM_SeccionF_Result>("REM_SeccionF", mesParameter, anioParameter);
        }
    }
}