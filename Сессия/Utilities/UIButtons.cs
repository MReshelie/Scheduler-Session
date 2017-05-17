using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraScheduler;
using System;
using System.Configuration;
using System.Linq;
using System.Runtime.CompilerServices;
using Сессия.Classes;

namespace Сессия
{
    public class UIButtons
    {
        #region Функционирование клавиш бокого меню
        static public bool UIButtonsEnabled(WindowsUIButtonPanel winUIBtPanel, bool[] blBtn)
        {
            for (int i = 0; i < winUIBtPanel.Buttons.Count; i++)

                if (winUIBtPanel.Buttons[i].GetType().Name != "DevExpress.XtraBars.Docking2010.WindowsUISeparator")
                    winUIBtPanel.Buttons[i].Properties.Enabled = blBtn[i];

            return true;
        }
        #endregion
    }

    public class Utilities
    {
        static public string[] GetConnectionStrings()
        {
            ConnectionStringSettingsCollection settings = ConfigurationManager.ConnectionStrings;

            int i = 0;
            string[] conStr = new string[settings.Count];

            Array.Clear(conStr, 0, settings.Count);

            if (settings != null)
            {
                foreach (ConnectionStringSettings cs in settings)
                {
                    //Console.WriteLine(cs.Name);
                    //Console.WriteLine(cs.ProviderName);
                    //Console.WriteLine(cs.ConnectionString);
                    conStr[i] = cs.Name;
                    i += 1;
                }
            }

            return conStr;
        }

        #region Трассировка вызова событий
        // Вызов трассировки:
        //
        //string[] tGrid = Utilities.TraceGrid();
        //Utilities.TraceMessage(tGrid[0].ToString());

        static public string[] TraceGrid([CallerMemberName()] string memberName = null)
        {
            return memberName.Split(new string[] { "_" }, StringSplitOptions.None);
        }

        static public void TraceMessage(string message,
        [System.Runtime.CompilerServices.CallerMemberName] string memberName = "",
        [System.Runtime.CompilerServices.CallerFilePath] string sourceFilePath = "",
        [System.Runtime.CompilerServices.CallerLineNumber] int sourceLineNumber = 0)
        {
            System.Diagnostics.Trace.WriteLine("message: " + message);
            System.Diagnostics.Trace.WriteLine("member name: " + memberName);
            System.Diagnostics.Trace.WriteLine("source file path: " + sourceFilePath);
            System.Diagnostics.Trace.WriteLine("source line number: " + sourceLineNumber);
        }
        #endregion
    }

    /*
     CREATE TABLE [dbo].[WorksheetData] (
    [IdRow] INT             IDENTITY (1, 1) NOT NULL,
    [Data]  VARBINARY (MAX) NULL,
    PRIMARY KEY CLUSTERED ([IdRow] ASC)
);

     */
}
