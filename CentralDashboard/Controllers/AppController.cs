using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Authentication;
using System.Web;
using System.Web.Mvc;

namespace CentralDashboard.Controllers
{
    public class AppController : Controller
    {
        protected Clases.ConnectionBuilder bdBuilder;
        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            base.OnActionExecuting(filterContext);
            if (Session["usuario"] == null)
            {
                throw new AuthenticationException();
            }
            bdBuilder = new Clases.ConnectionBuilder(Session);
        }

        protected override void OnException(ExceptionContext filterContext)
        {
            if (filterContext.Exception.GetType() == typeof(AuthenticationException))
            {
                filterContext.ExceptionHandled = true;
                filterContext.Result = this.RedirectToAction("Index", "Home", new { area = "" });
            }

            base.OnException(filterContext);
        }
        protected string GetUsuario()
        {
            return (string)Session["usuario"];
        }
    }
}