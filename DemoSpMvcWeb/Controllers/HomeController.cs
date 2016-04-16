using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DemoSpMvcWeb.Models;

namespace DemoSpMvcWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;

                    clientContext.Load(spUser, user => user.Title);

                    clientContext.ExecuteQuery();

                    ViewBag.UserName = spUser.Title;
                }
            }

            return View();
        }

        public ActionResult ListaPersonas()
        {
            var model = new List<Persona>();
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var cliente = spContext.CreateUserClientContextForSPHost())
            {
                if (cliente != null)
                {
                    var web = cliente.Web;
                    cliente.Load(web);
                    cliente.ExecuteQuery();

                    var listas = web.Lists;
                    cliente.Load(listas);
                    cliente.ExecuteQuery();

                    var personas = listas.GetByTitle("Personas");
                    cliente.Load(personas);
                    cliente.ExecuteQuery();

                    var query = new CamlQuery();
                    var listadoPersonas = personas.GetItems(query);
                    cliente.Load(listadoPersonas);
                    cliente.ExecuteQuery();

                    foreach (var per in listadoPersonas)
                    {
                        model.Add(new Persona()
                        {
                            Apellido = per["Title"].ToString(),
                            Nombre = per["Nombre"].ToString(),
                            Ciudad = per["Ciudad"].ToString()
                        });
                    }
                }
                return View(model);
            }
        }

    }
}
