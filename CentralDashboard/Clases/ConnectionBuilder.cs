using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Web;

namespace CentralDashboard.Clases
{

    public class ConnectionBuilder
    {
        private string usuario;
        private string pass;
        private string servidor;
        private string[] configuracionEnti = { 
            "metadata=res://*/Models.EntiCorporativa.ModeloEnti.csdl|res://*/Models.EntiCorporativa.ModeloEnti.ssdl|res://*/Models.EntiCorporativa.ModeloEnti.msl;provider=System.Data.SqlClient;provider connection string=\"data source=",
            "",
            ";initial catalog=BD_ENTI_CORPORATIVA;user id=", 
            "", 
            ";password=", 
            "", 
            ";MultipleActiveResultSets=True;App=Dashboard\"" };
        private string[] configuracionAbastecimiento =
        {
            "metadata=res://*/Models.Abastecimiento.Abastecimiento.csdl|res://*/Models.Abastecimiento.Abastecimiento.ssdl|res://*/Models.Abastecimiento.Abastecimiento.msl;provider=System.Data.SqlClient;provider connection string=\"data source=",
            "",
            ";initial catalog=BD_ABASTECIMIENTO;user id=",
            "",
            ";password=",
            "",
            ";MultipleActiveResultSets=True;App=Dashboard;Connect Timeout=60\"" };

        public ConnectionBuilder(HttpSessionStateBase session)
        {
            usuario = (string)session["usuario"];
            pass = (string)session["pass"];
            servidor = (string)session["servidor"];
        }

        public ConnectionBuilder(string usuario, string pass, string servidor)
        {
            this.usuario = usuario;
            this.pass = pass;
            this.servidor = servidor;
        }

        private string GenerarStringBuilder(string[] configuracion)
        {
            var strBuilder = new StringBuilder("");
            strBuilder.Append(configuracion[0]);
            strBuilder.Append(servidor);
            strBuilder.Append(configuracion[2]);
            strBuilder.Append(usuario);
            strBuilder.Append(configuracion[4]);
            strBuilder.Append(pass);
            strBuilder.Append(configuracion[6]);
            return strBuilder.ToString();
        }

        public Models.EntiCorporativa.Entities GetEntiCorporativa()
        {
            return new Models.EntiCorporativa.Entities(GenerarStringBuilder(configuracionEnti));
        }

        public Models.Abastecimiento.BD_ABASTECIMIENTOEntities1 GetAbastecimiento()
        {
            return new Models.Abastecimiento.BD_ABASTECIMIENTOEntities1(GenerarStringBuilder(configuracionAbastecimiento));
        }
    }
}

namespace CentralDashboard.Models.EntiCorporativa
{
    public partial class Entities : DbContext
    {
        public Entities(string conexion)
            : base(conexion)
        {
        }
    }
}

namespace CentralDashboard.Models.Abastecimiento
{
    public partial class BD_ABASTECIMIENTOEntities1 : DbContext
    {
        public BD_ABASTECIMIENTOEntities1(string conexion)
            : base(conexion)
        {
        }
    }
}