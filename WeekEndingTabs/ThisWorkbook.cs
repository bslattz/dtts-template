using System;
using System.Runtime.InteropServices;
using WeekEnding;

namespace WeekEndingTabs
{
    [ComVisible(true)]
    public interface IWeekending
    {
        string DisplayTaggedSheets ();
        string DisplayDates ();
        void Refresh();
        void Log(string message);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class QueryWeekending : IWeekending
    {
        private readonly WeekEnding.WeekEnding _we;
        public WeekEnding.WeekEnding We =>
            _we ?? Globals.ThisWorkbook?.WeekEnding;

        public QueryWeekending (WeekEnding.WeekEnding we)
        {
            _we = we;
            We?.Log("QueryWeekending: Startup");
        }

        ~QueryWeekending()
        {
            We?.Log("QueryWeekending: Shutdown");
        }
        string IWeekending.DisplayTaggedSheets ()
        {
            return We.DisplayTaggedSheets();
        }
        string IWeekending.DisplayDates ()
        {
            return We.DisplayDates();
        }

        void IWeekending.Refresh ()
        {
            We.Refresh();
        }

        void IWeekending.Log (string message)
        {
            We?.Log(message);
        }
    }

    public partial class ThisWorkbook
    {
        public WeekEnding.WeekEnding WeekEnding;
        private QueryWeekending _qwe;

        private void ThisWorkbook_Startup (object sender, System.EventArgs e)
        {
            WeekEnding = new WeekEnding.WeekEnding(this, Globals.Factory);
        }

        private void ThisWorkbook_Shutdown (object sender, System.EventArgs e)
        {
        }

        protected override object GetAutomationObject ()
        {
            return _qwe ?? (_qwe = new QueryWeekending(WeekEnding));
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup ()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion
    }
}
