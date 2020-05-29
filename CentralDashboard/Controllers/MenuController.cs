using CentralDashboard.Models.EntiCorporativa;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CentralDashboard.Controllers
{
    public class MenuController : AppController
    {
        [ChildActionOnly]
        // GET: Menu
        public ActionResult Superior() 
        {
            return View();
        }

        [ChildActionOnly]
        public PartialViewResult Izquierda()
        {
            string idUsuario = GetUsuario();
            var bd = bdBuilder.GetEntiCorporativa();
            var paginas = bd.USR_PaginaSitioWeb.Where(x => x.Funcion == "INDEX").ToList();
            var listado = new List<USR_PaginaSitioWeb>();
            foreach (var pagina in paginas)
            {
                if (pagina.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario))
                {
                    listado.Add(pagina);
                }
            }
            ViewBag.Paginas = listado;
            return PartialView();
        }

        [ChildActionOnly]
        public PartialViewResult IzquierdaItem(int idItem)
        {
            string idUsuario = GetUsuario();
            var bd = bdBuilder.GetEntiCorporativa();
            var paginaPpal = bd.USR_PaginaSitioWeb.First(x => x.Id == idItem);
            ViewBag.PaginaPpal = paginaPpal;
            ViewBag.PaginasHijas = bd.USR_PaginaSitioWeb.Where(x => x.Controlador == paginaPpal.Controlador && x.Funcion != "INDEX").ToList();
            return PartialView();
        }
    }
}