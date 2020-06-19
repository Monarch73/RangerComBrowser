using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using WaitHelpers = SeleniumExtras.WaitHelpers;

namespace RangerComBrowser
{
#if BIT32
    [Guid("0CB1F887-683C-4AAF-B5A1-372246A8F348")]
#else
    [Guid("F466DD36-353C-47FA-8727-42EF8442F887")]
#endif
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
        /// <param name="sendEnter">send enter after the user input.</param>
        /// <returns></returns>
        bool SendKeysToId(string id, string keys, bool sendEnter);

        /// <summary>
        /// Emulates user input to control identified by Name.
        /// </summary>
        /// <param name="name">Name of the element.</param>
        /// <param name="keys">Userinput.</param>
        /// <param name="sendEnter">send enter after the user input.</param>
        /// <returns></returns>
        bool SendKeysToName(string name, string keys, bool sendEnter);

        /// <summary>
        /// Emulates user input to control identified by xpath.
        /// </summary>
        /// <param name="xpath">XPath of the element.</param>
        /// <param name="keys">User input.</param>
        /// <param name="sendEnter">send enter after the user input.</param>
        /// <returns></returns>
        bool SendKeysToXPath(string xpath, string keys, bool sendEnter);

        /// <summary>
        /// Simulates a mouseclick on an element identified by xpath.
        /// </summary>
        /// <param name="xpath">XPath of the element.</param>
        /// <returns></returns>
        bool SendClickToXPath(string xpath);

        /// <summary>
        /// Simulates a mouseclick on an element identified by Name.
        /// </summary>
        /// <param name="name">Name of the element.</param>
        /// <returns></returns>
        bool SendClickToName(string name);

        /// <summary>
        /// Simulates a mouseclick on an element identified by Id.
        /// </summary>
        /// <param name="id">Id of the element to send the mouseclick to.</param>
        /// <returns></returns>
        bool SendClickToId(string id);

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

        /// <summary>
        /// Switch to Window Name
        /// </summary>
        /// <param name="windowName">Name of the window to switch to.</param>
        /// <returns>true on success.</returns>
        bool SwitchTo(string windowName);

        /// <summary>
        /// Gets a list of window names.
        /// </summary>
        /// <returns>Names of all windows.</returns>
        string[] GetWindowNames();

        /// <summary>
        /// Gets the current window name.
        /// </summary>
        /// <returns>Name of the current window.</returns>
        string GetWindowName();

        /// <summary>
        /// Gets the windows title of the current focused window.
        /// </summary>
        /// <returns>Title of the windows.</returns>
        string GetWindowTitle();

        /// <summary>
        /// Switches the focus to a window identified by title.
        /// </summary>
        /// <param name="title">Title of the window to switch to.</param>
        /// <returns></returns>
        bool SwitchToTitle(string title);

        /// <summary>
        /// Check if an xpath exists.
        /// </summary>
        /// <param name="xpath">xpath to look for.</param>
        /// <returns>true, if xpath is found</returns>
        bool CheckXPath(string xpath);

        /// <summary>
        /// Check if an id exists.
        /// </summary>
        /// <param name="id">id to look for.</param>
        /// <returns>true, if id is found.</returns>
        bool CheckId(string id);

        /// <summary>
        /// Check if an name exists.
        /// </summary>
        /// <param name="name">name to look for.</param>
        /// <returns>true, if name is found.</returns>
        bool CheckName(string name);

        WebItem[] GetItemsFromSelectByXPath(string xpath);

        WebItem[] GetItemsFromSelectByName(string name);

        WebItem[] GetItemsFromSelectById(string id);


    }
}
