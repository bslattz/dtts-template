using System.Runtime.InteropServices;
using WeekEnding;

namespace WeekEndingTabs
{
    [ComVisible(true)]
    public interface IWeekending
    {
        string DisplayTaggedSheets ();
        string DisplayDates ();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class QueryWeekending : IWeekending
    {
        private readonly WeekEnding.WeekEnding _we;

        public WeekEnding.WeekEnding We =>
            _we == null ? Globals.ThisWorkbook?.WeekEnding : _we;

        public QueryWeekending ()
        {
            _we = Globals.ThisWorkbook?.WeekEnding;
        }
        string IWeekending.DisplayTaggedSheets ()
        {
            return We.DisplayTaggedSheets();
        }
        string IWeekending.DisplayDates ()
        {
            return We.DisplayDates();
        }
    }

    public partial class ThisWorkbook
    {
        public WeekEnding.WeekEnding WeekEnding;

        private void ThisWorkbook_Startup (object sender, System.EventArgs e)
        {
            WeekEnding = new WeekEnding.WeekEnding(this, Globals.Factory);
        }

        private void ThisWorkbook_Shutdown (object sender, System.EventArgs e)
        {
        }

        protected override object GetAutomationObject ()
        {
            return new QueryWeekending();
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
