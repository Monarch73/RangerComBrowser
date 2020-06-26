using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using WaitHelpers = SeleniumExtras.WaitHelpers;
using Newtonsoft.Json;

namespace RangerComBrowser
{
#if BIT32
    [Guid("2C6E9D05-80EE-4B6E-8890-DADA381BDE45"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(ITypeLib))]
    [ProgId("RangerComBrowser32.TypeLib")]
#else
    [Guid("43EED648-88B4-440E-9026-82004277CB42"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(ITypeLib))]
    [ProgId("RangerComBrowser.TypeLib")]
#endif
    [ComVisible(true)]
    public class TypeLib : ITypeLib, IDisposable
    {
        private ChromeDriver driver;


        public bool LoadBrowser()
        {
            if (this.driver == null)
            {
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                this.driver = new ChromeDriver(chromeDriverService);
                return true;
            }
            return false;
        }

        public bool UnloadBrowser()
        {
            if (this.driver != null)
            {
                driver.Dispose();
                driver = null;
                return true;
            }

            return false;
        }

        public bool GoToUrl(string url)
        {
            if (this.driver != null)
            {
                this.driver.Navigate().GoToUrl(url);
                return true;
            }

            return false;
        }

        public bool SendKeysToId(string id, string keys, bool sendEnter)
        {
            if (this.driver != null)
            {
                var element = this.driver.FindElementById(id);//.SendKeys(keys);
                element.SendKeys(keys);
                if (sendEnter)
                {
                    element.SendKeys(Keys.Enter);
                }
                return true;
            }

            return false;
        }

        public bool SendKeysToName(string name, string keys, bool sendEnter)
        {
            if (this.driver != null)
            {
                var element = this.driver.FindElementByName(name); //.SendKeys(keys);
                element.SendKeys(keys);
                if (sendEnter)
                {
                    element.SendKeys(Keys.Enter);
                }
                return true;
            }

            return false;
        }

        public bool SendKeysToXPath(string xpath, string keys, bool sendEnter)
        {
            if (this.driver != null)
            {
                var element = this.driver.FindElementByXPath(xpath);//.SendKeys(keys);
                element.SendKeys(keys);
                if (sendEnter)
                {
                    element.SendKeys(Keys.Enter);
                }
                return true;
            }

            return false;
        }
        public bool SendClickToXPath(string xpath)
        {
            if (this.driver != null)
            {
                this.driver.FindElementByXPath(xpath).Click();
                return true;
            }

            return false;

        }

        public bool SendClickToName(string name)
        {
            if (this.driver != null)
            {
                this.driver.FindElementByName(name).Click();
                return true;
            }

            return false;
        }

        public bool SendClickToId(string id)
        {
            if (this.driver != null)
            {
                this.driver.FindElementById(id).Click();
                return true;
            }

            return false;
        }

        public string GetTextFromId(string id)
        {
            return this.driver?.FindElementById(id)?.Text;
        }

        public string GetTextFromName(string name)
        {
            return this.driver.FindElementByName(name).Text;
        }

        public string GetTextFromXPath(string xpath)
        {
            return this.driver?.FindElementByXPath(xpath)?.Text;
        }

        public object ExecuteScript(string script, params object[] args)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            return js?.ExecuteScript("return document.title", args);
        }

        public bool WaitForJavascript(string script, string expectedResult)
        {
            if (this.driver != null)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(1, 0, 10));
                wait.Until(wd => js.ExecuteScript(script).ToString() == expectedResult);
                return true;
            }

            return false;
        }

        public bool WaitForIdAvail(string id)
        {
            if (this.driver != null)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(1, 0, 10));
                wait.Until(WaitHelpers.ExpectedConditions.ElementExists(By.Id(id)));

                return true;
            }

            return false;
        }

        public bool WaitForNameAvail(string name)
        {
            if (this.driver != null)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(1, 0, 10));
                wait.Until(WaitHelpers.ExpectedConditions.ElementExists(By.Name(name)));

                return true;
            }

            return false;
        }

        public bool WaitForXPathAvail(string xpath)
        {
            if (this.driver != null)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(1, 0, 10));
                wait.Until(WaitHelpers.ExpectedConditions.ElementExists(By.XPath(xpath)));

                return true;
            }

            return false;
        }

        public bool WaitForUrl(string url)
        {
            if (this.driver != null)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(1, 0, 10));
                wait.Until(WaitHelpers.ExpectedConditions.UrlMatches(url));

                return true;
            }

            return false;
        }

        public bool SwitchTo(string windowName)
        {
            if (this.driver != null)
            {
                this.driver.SwitchTo().Window(windowName);
                return true;
            }
            return false;
        }

        public string[] GetWindowNames()
        {
            if (this.driver != null)
            {
                return this.driver.WindowHandles.ToArray<string>();
            }
            return null;
        }

        public string GetWindowName()
        {
            if(this.driver != null)
            {
                return this.driver.CurrentWindowHandle;
            }

            return null;
        }

        public string GetWindowTitle()
        {
            if (this.driver != null)
            {
                return this.driver.Title;
            }

            return null;
        }

        public bool SwitchToTitle(string title)
        {
            if (this.driver != null)
            {
                var name = this.GetWindowTitle();
                if (title != name)
                {
                    var names = this.GetWindowNames();
                    foreach (var handle in names)
                    {
                        this.SwitchTo(handle);
                        name = this.GetWindowTitle();
                        if (name == title)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public bool CheckXPath(string xpath) => this.Check(By.XPath(xpath));

        public bool CheckId(string id) => this.Check(By.Id(id));

        public bool CheckName(string name) => this.Check(By.Name(name));

        private bool Check(By by)
        {
            if (this.driver != null)
            {
                return this.driver.FindElements(by).Count != 0;
            }

            return false;
        }

        public string GetItemsFromSelectByXPath(string xpath) => this.GetItemsFromSelect(By.XPath(xpath));
        public string GetItemsFromSelectByName(string name) => this.GetItemsFromSelect(By.Name(name)); 
        public string GetItemsFromSelectById(string id) => this.GetItemsFromSelect(By.Id(id));

        private string GetItemsFromSelect(By by)
        {
            if (this.driver != null)
            {
                var selectElement = this.driver.FindElement(by);
                var select = new SelectElement(selectElement);
                var options = new List<WebItem>();
                int index = 0;
                foreach (var option in select.Options)
                {
                    options.Add(new WebItem(index++, option.Text, option.GetAttribute("value")));
                }

                return JsonConvert.SerializeObject(options.ToArray());
            }
            return null;
        }

        public bool SelectOptionIndexByXpath(string xpath, int index) => this.SelectOptionIndex(By.XPath(xpath), index);
        public bool SelectOptionIndexById(string id, int index) => this.SelectOptionIndex(By.Id(id), index);
        public bool SelectOptionIndexByName(string name, int index) => this.SelectOptionIndex(By.Name(name), index);

        public bool SelectOptionValueByXpath(string xpath, string value) => this.SelectOptionValue(By.XPath(xpath), value);
        public bool SelectOptionValueById(string id, string value) => this.SelectOptionValue(By.Id(id), value);
        public bool SelectOptionValueByName(string name, string value) => this.SelectOptionValue(By.Name(name), value);

        public bool SelectOptionTextByXpath(string xpath, string text) => this.SelectOptionText(By.XPath(xpath), text);
        public bool SelectOptionTextById(string id, string text) => this.SelectOptionText(By.Id(id), text);
        public bool SelectOptionTextByName(string name, string text) => this.SelectOptionText(By.Name(name), text);

        private bool SelectOptionText(By by, string text)
        {
            if (this.driver != null)
            {
                var selectElement = this.driver.FindElement(by);
                var select = new SelectElement(selectElement);
                select.SelectByText(text);
                return true;
            }

            return false;
        }

        private bool SelectOptionIndex(By by, int index)
        {
            if (this.driver != null)
            {
                var selectElement = this.driver.FindElement(by);
                var select = new SelectElement(selectElement);
                select.SelectByIndex(index);
                return true;
            }

            return false;
        }

        private bool SelectOptionValue(By by, string value)
        {
            if (this.driver != null)
            {
                var selectElement = this.driver.FindElement(by);
                var select = new SelectElement(selectElement);
                select.SelectByValue(value);
                return true;
            }

            return false;

        }


        public int GetVersion()
        {
            return 100;
        }

#region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    this.UnloadBrowser();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~Interface()
        // {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
#endregion

    }

}
