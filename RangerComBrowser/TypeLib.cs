using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WaitHelpers = SeleniumExtras.WaitHelpers;


namespace RangerComBrowser
{
    [Guid("0CB1F887-683C-4AAF-B5A1-372246A8F348")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITypeLib
    {
        /// <summary>
        /// Gets the Version Number
        /// </summary>
        /// <returns>Version times 100.</returns>
        int GetVersion();

        /// <summary>
        /// Loads and displays the browser window.
        /// </summary>
        /// <returns>true if success, false if  browser is already loaded.</returns>
        bool LoadBrowser();

        /// <summary>
        /// Unloads and hides the browser window.
        /// </summary>
        /// <returns>true, if browser was successfully unloaded.</returns>
        bool UnloadBrowser();

        /// <summary>
        /// Requests browser to load url.
        /// </summary>
        /// <param name="url"></param>
        /// <returns>true, if browser was loaded.</returns>
        bool GoToUrl(string url);

        /// <summary>
        /// Emulates user input to control identified by Id.
        /// </summary>
        /// <param name="id">Id of the element.</param>
        /// <param name="keys">userinput.</param>
        /// <returns></returns>
        bool SendKeysToId(string id, string keys);

        /// <summary>
        /// Emulates user input to control identified by Name.
        /// </summary>
        /// <param name="name">Name of the element.</param>
        /// <param name="keys">Userinput.</param>
        /// <returns></returns>
        bool SendKeysToName(string name, string keys);

        /// <summary>
        /// Emulates user input to control identified by xpath.
        /// </summary>
        /// <param name="xpath">XPath of the element.</param>
        /// <param name="keys">User input.</param>
        /// <returns></returns>
        bool SendKeysToXPath(string xpath, string keys);

        /// <summary>
        /// Gets the Text-Property of element identified by Id.
        /// </summary>
        /// <param name="id">Id of the element.</param>
        /// <returns>Text of the element, null if element not found or browser not loaded.</returns>
        string GetTextFromId(string id);

        /// <summary>
        /// Gets the text-property of the element identified by Name.
        /// </summary>
        /// <param name="name">Name of the element.</param>
        /// <returns>Text of the element, null if element not found or browser not loaded.</returns>
        string GetTextFromName(string name);

        /// <summary>
        /// Gets the text-property of the element identified by xpath.
        /// </summary>
        /// <param name="xpath">xpath of the element.</param>
        /// <returns>Text of the element, null if element not found or browser not loaded.</returns>
        string GetTextFromXPath(string xpath);

        /// <summary>
        /// Executes javascript in browser.
        /// </summary>
        /// <param name="script">Actual script.</param>
        /// <param name="args">Arguments. Can be null if not used.</param>
        /// <returns>Result of the javascript, null if no browser was loaded.</returns>
        object ExecuteScript(string script, params object[] args);

        /// <summary>
        /// Waits and blocks execution until provided javascript returns specific result.
        /// </summary>
        /// <param name="script">javascript to execute.</param>
        /// <param name="expectedResult">Expected result of the javascript.</param>
        /// <returns></returns>
        bool WaitForJavascript(string script, string expectedResult);

        /// <summary>
        /// Waits and blocks execution until element identified by id becomes available.
        /// </summary>
        /// <param name="id">Id of the element.</param>
        /// <returns>true, if successfull, false on any error.</returns>
        bool WaitForIdAvail(string id);

        /// <summary>
        /// Waits and blocks execution until element identified by Name becomes available.
        /// </summary>
        /// <param name="name">Name of the element.</param>
        /// <returns>true, if successfull, false on any error.</returns>
        bool WaitForNameAvail(string name);

        /// <summary>
        /// Waits and blocks execution until element identified by xpath becomes available.
        /// </summary>
        /// <param name="xpath">xpath of the element.</param>
        /// <returns>true, if successfull, false on any error.</returns>
        bool WaitForXPathAvail(string xpath);

        /// <summary>
        /// Waits and blocks execution until browser url matches regex expression
        /// </summary>
        /// <param name="url">Regex expression to be matched against browser url.</param>
        /// <returns>true, if successfull, false on any error.</returns>
        bool WaitForUrl(string url);
    }

    [Guid("2C6E9D05-80EE-4B6E-8890-DADA381BDE45"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(ITypeLib))]
    [ComVisible(true)]
    [ProgId("RangerComBrowser.TypeLib")]
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

        public bool SendKeysToId(string id, string keys)
        {
            if (this.driver != null)
            {
                this.driver.FindElementById(id).SendKeys(keys);
                return true;
            }

            return false;
        }

        public bool SendKeysToName(string name, string keys)
        {
            if (this.driver != null)
            {
                this.driver.FindElementByName(name).SendKeys(keys);
                return true;
            }

            return false;
        }

        public bool SendKeysToXPath(string xpath, string keys)
        {
            if (this.driver != null)
            {
                this.driver.FindElementByXPath(xpath).SendKeys(keys);
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
