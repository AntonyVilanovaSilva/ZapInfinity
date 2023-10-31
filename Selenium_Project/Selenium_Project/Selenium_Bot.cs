using EasyAutomationFramework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Selenium_Project
{
    public class Selenium_Bot:Web
    {
        public void EnviarMensagem(string mensagem)
        {
            StartBrowser(TypeDriver.GoogleChorme);

            Navigate("https://www.google.com/");

            WaitForLoad();

            var elementSearch = AssignValue(TypeElement.Xpath, "//*[@id=\"APjFqb\"]", mensagem);

            elementSearch.element.SendKeys(OpenQA.Selenium.Keys.Enter);

        }
    }
}
