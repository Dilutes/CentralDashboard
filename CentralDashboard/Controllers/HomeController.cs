using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CentralDashboard.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.FocoUser = "autofocus";
            return View();
        }

        [HttpPost]
        public ActionResult Index(string usuario, string pass)
        {
            var bd = (new Clases.ConnectionBuilder(usuario, pass, "10.67.84.215")).GetEntiCorporativa();
            if(bd.Database.Exists()){
                if(bd.USR_PermisoSitioWeb.Any(x=>x.Usuario == usuario))
                {
                    Session["usuario"] = usuario;
                    Session["pass"] = pass;
                    Session["servidor"] = "10.67.84.215";
                    return Redirect("/Administracion/");
                }
                else
                {
                    ViewBag.Error = "No tiene privilegios para usar este sistema";
                }
            }
            else
            {
                ViewBag.Error = "Usuario y/o contraseña incorrecta";
            }
            ViewBag.Usuario = usuario;
            ViewBag.FocoPass = "autofocus";
            
            return View();
        }

        public ActionResult CerrarSesion()
        {
            Session.Clear();
            return Redirect("/");
        }
    }
}